import argparse
from contextlib import closing
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
import json
import os
from pathlib import Path, PureWindowsPath
import re
import sqlite3
import subprocess
import sys
from threading import Lock, Thread
import tomllib
from urllib.parse import parse_qs, quote, urlsplit
from uuid import uuid4

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile

from catalog import References
from robot import IRKUTSK, MAX_BYTES


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "web_assets"
STATIC = {"/": ("index.html", "text/html; charset=utf-8"),
          "/styles.css": ("styles.css", "text/css; charset=utf-8"),
          "/app.js": ("app.js", "text/javascript; charset=utf-8")}
EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def filename_ok(value):
    return (0 < len(value) <= 200 and not value.startswith((".", "~$"))
            and not value.endswith((" ", ".")) and not re.search(r'[<>:"/\\|?*\x00-\x1f]', value)
            and not PureWindowsPath(value).is_reserved())


class Panel:
    def __init__(self, config):
        self.config = config
        paths = tomllib.loads(config.read_text(encoding="utf-8"))["paths"]
        self.exchange = config.parent / paths["exchange"]
        self.reference = config.parent / paths["references"]
        self.state = config.parent / paths["state"]
        self.inbox = self.exchange / "Входящие файлы"
        self.reviewed = self.exchange / "Проверенные"
        self.outputs = {"report": self.exchange / "Отчёты", "upload": self.exchange / "Загрузка"}
        self.guard = Lock()
        self.run = None
        self.reference_version = None
        self.profiles = {}
        self.reference_issue = ""

    def profile_info(self):
        try:
            info = self.reference.stat()
        except FileNotFoundError:
            return {}, "Добавьте книгу справочников."
        version = (info.st_mtime_ns, info.st_size)
        with self.guard:
            if version != self.reference_version:
                try:
                    with closing(load_workbook(BytesIO(self.reference.read_bytes()), keep_links=False)) as book:
                        ref = References(book)
                    self.profiles = {supplier: profile[0] for supplier, profile in sorted(ref.profiles.items())
                                     if filename_ok(supplier)}
                    self.reference_issue = ""
                except (InvalidFileException, BadZipFile, KeyError, SyntaxError, TypeError, ValueError) as error:
                    self.profiles = {}
                    self.reference_issue = f"Проверьте книгу справочников. {error}"
                self.reference_version = version
            return self.profiles, self.reference_issue

    def query(self, sql, args=()):
        try:
            with closing(sqlite3.connect(self.state.resolve().as_uri() + "?mode=ro", uri=True, timeout=2)) as db:
                db.row_factory = sqlite3.Row
                return [dict(row) for row in db.execute(sql, args)]
        except sqlite3.OperationalError as error:
            if not self.state.exists() or str(error).startswith("no such table:"):
                return []
            raise

    def output_path(self, kind, name):
        directory = self.outputs[kind].resolve()
        if not filename_ok(name) or Path(name).suffix.lower() != ".xlsx":
            return None
        path = (directory / name).resolve()
        return path if path.parent == directory and path.is_file() else None

    def snapshot(self):
        profiles, issue = self.profile_info()
        files = self.query("SELECT id,supplier,name,status,counts,report FROM files ORDER BY rowid DESC LIMIT 200")
        for file in files:
            name = file.pop("report")
            file["counts"] = json.loads(file["counts"])
            file["downloads"] = [kind for kind in self.outputs if self.output_path(kind, name)]
        journal = self.query("SELECT at,stage,status,code FROM journal ORDER BY rowid DESC LIMIT 30")
        incoming = [dict(supplier=path.parent.name, name=path.name)
                    for path in sorted(self.inbox.glob("*/*")) if path.is_file() and filename_ok(path.name)]
        waiting = [path.name for path in sorted(self.reviewed.glob("*.xlsx"))
                   if path.is_file() and filename_ok(path.name)]
        with self.guard:
            result = dict(self.run) if self.run else None
        return dict(profiles=profiles, issue=issue, files=files, journal=journal,
                    incoming=incoming, waiting=waiting, run=result)

    def start(self):
        _, issue = self.profile_info()
        if issue:
            raise ValueError(issue)
        with self.guard:
            if self.run and self.run["running"]:
                return False
            self.run = dict(running=True, at=datetime.now(IRKUTSK).isoformat(), code=None, output="")
            Thread(target=self.run_cycle, daemon=False).start()
        return True

    def run_cycle(self):
        try:
            result = subprocess.run([sys.executable, "-u", str(ROOT / "robot.py"),
                                     "--config", str(self.config)],
                                    cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                    encoding="utf-8", errors="replace", check=False,
                                    env={**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"})
            code, output = result.returncode, result.stdout.strip()
        except OSError:
            code, output = 1, "Не удалось запустить робота. Проверьте установку Python."
        with self.guard:
            self.run.update(running=False, code=code, output=output[-3000:])

    def upload(self, kind, supplier, name, body):
        if not filename_ok(name):
            raise ValueError("Проверьте имя файла.")
        ext = Path(name).suffix.lower().lstrip(".")
        if kind == "review":
            if ext != "xlsx":
                raise ValueError("Проверенный отчёт должен быть в формате XLSX.")
            directory = self.reviewed
        else:
            profiles, issue = self.profile_info()
            if issue:
                raise ValueError(issue)
            if supplier not in profiles or profiles[supplier] != ext:
                raise ValueError("Выберите поставщика и файл в формате, указанном в его профиле.")
            directory = self.inbox / supplier
            if directory.resolve() != self.inbox.resolve() / supplier:
                raise ValueError("Папка поставщика недоступна.")
        directory.mkdir(parents=True, exist_ok=True)
        destination = directory / name
        temp = directory / f".{uuid4().hex}.part"
        try:
            temp.write_bytes(body)
            if os.name == "nt":
                temp.rename(destination)
            else:
                os.link(temp, destination)
        finally:
            temp.unlink(missing_ok=True)


class Handler(BaseHTTPRequestHandler):
    def log_request(self, code="-", size="-"):
        if self.path != "/api/status" or code != HTTPStatus.OK:
            super().log_request(code, size)

    def send_data(self, code, content, content_type, attachment=None):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy", "default-src 'self'; frame-ancestors 'none'; form-action 'self'")
        if attachment:
            self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + quote(attachment, safe=""))
        self.end_headers()
        self.wfile.write(content)

    def json(self, code, data):
        self.send_data(code, json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def dispatch(self, write=False):
        port = self.server.server_port
        if self.headers.get("Host") not in {f"127.0.0.1:{port}", f"localhost:{port}"}:
            self.json(HTTPStatus.FORBIDDEN, {"error": "Панель доступна на этом компьютере."})
            return
        target = urlsplit(self.path)
        params = {key: values[0] for key, values in parse_qs(target.query).items()}
        app = self.server.panel
        try:
            if write:
                if self.headers.get("X-DomTorg") != "1":
                    self.json(HTTPStatus.FORBIDDEN, {"error": "Откройте панель и повторите действие."})
                elif target.path == "/api/run":
                    started = app.start()
                    self.json(HTTPStatus.ACCEPTED if started else HTTPStatus.CONFLICT,
                              {"ok": True} if started else {"error": "Цикл уже выполняется."})
                elif target.path == "/api/upload":
                    self.receive_upload(app, params)
                else:
                    self.json(HTTPStatus.NOT_FOUND, {"error": "Действие не найдено."})
            elif target.path in {"/help", "/help/"}:
                filename, content_type = STATIC["/"]
                page = (ASSETS / filename).read_text(encoding="utf-8")
                help_block = (ASSETS / "help.html").read_text(encoding="utf-8")
                page = page.replace("ДомТорг – проверка робота", "ДомТорг – справка по панели", 1)
                page = page.replace("<!-- panel-help -->", help_block, 1)
                self.send_data(HTTPStatus.OK, page.encode("utf-8"), content_type)
            elif target.path in STATIC:
                filename, content_type = STATIC[target.path]
                self.send_data(HTTPStatus.OK, (ASSETS / filename).read_bytes(), content_type)
            elif target.path == "/api/status":
                self.json(HTTPStatus.OK, app.snapshot())
            elif target.path == "/download" and params.get("kind") in app.outputs:
                files = app.query("SELECT report FROM files WHERE id=?", (params.get("id", ""),))
                path = app.output_path(params["kind"], files[0]["report"]) if files else None
                if path:
                    self.send_data(HTTPStatus.OK, path.read_bytes(), EXCEL, path.name)
                else:
                    self.json(HTTPStatus.NOT_FOUND, {"error": "Файл больше недоступен."})
            else:
                self.json(HTTPStatus.NOT_FOUND, {"error": "Страница не найдена."})
        except ValueError as error:
            self.json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
        except FileExistsError:
            self.json(HTTPStatus.CONFLICT, {"error": "Файл с таким именем уже ждёт обработки."})
        except (BrokenPipeError, ConnectionResetError):
            return
        except (OSError, sqlite3.Error):
            self.json(HTTPStatus.SERVICE_UNAVAILABLE, {"error": "Не удалось прочитать или записать данные. Повторите попытку."})

    def receive_upload(self, app, params):
        kind = params.get("kind")
        size = int(self.headers.get("Content-Length", "0"))
        if kind not in {"supplier", "review"} or self.headers.get_content_type() != "application/octet-stream":
            raise ValueError("Выберите файл для загрузки.")
        if not 0 < size <= MAX_BYTES:
            self.json(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, {"error": "Размер файла должен быть от 1 байта до 10 МБ."})
            return
        self.connection.settimeout(15)
        body = self.rfile.read(size)
        if len(body) != size:
            raise ValueError("Загрузка прервалась. Повторите попытку.")
        app.upload(kind, params.get("supplier", ""), params.get("name", ""), body)
        self.json(HTTPStatus.CREATED, {"ok": True})

    def do_GET(self):
        self.dispatch()

    def do_POST(self):
        self.dispatch(write=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Локальная панель ДомТорг")
    parser.add_argument("--config", type=Path, default=ROOT / "config.toml")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    with ThreadingHTTPServer(("127.0.0.1", args.port), Handler) as server:
        server.panel = Panel(args.config.resolve())
        print(f"ДомТорг: http://127.0.0.1:{server.server_port}/", flush=True)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
