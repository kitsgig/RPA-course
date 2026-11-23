import argparse
from collections import Counter
from contextlib import closing, contextmanager
from datetime import datetime, time, timedelta
from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path
import smtplib
import sqlite3
import ssl
import sys
from time import sleep
import tomllib
from email.message import EmailMessage
from zipfile import BadZipFile
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from catalog import (CHECK, FORMATS, InputError, NEW, ProfileError, References,
                     attributes, categorize, clean, factor, number, parts, read_positions,
                     validate_result)
from credentials import read_windows_credential
from gigachat import GigaChatClient, GigaChatConfig
from workbooks import DECISIONS, REPORT, report_book, save_book, save_links, upload_book


REGISTERED = "Зарегистрирован"
SETUP = "Требует настройки"
REPORTED = "Отчёт сформирован"
PARTIAL = "Частично обработан"
REVIEWED = "Решения учтены"
ERROR = "Ошибка"
MAX_BYTES = 10 * 1024 * 1024
IRKUTSK = ZoneInfo("Asia/Irkutsk")
NETWORK_ERRORS = {51, 53, 54, 55, 59, 64, 67, 121, 1231, 1232}
DATABASE = """
CREATE TABLE IF NOT EXISTS files (
    id TEXT PRIMARY KEY, supplier TEXT NOT NULL, name TEXT NOT NULL,
    status TEXT NOT NULL, version TEXT NOT NULL, counts TEXT NOT NULL DEFAULT '{}',
    report TEXT NOT NULL, archived INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS positions (
    file_id TEXT NOT NULL REFERENCES files(id), row_no INTEGER NOT NULL,
    data TEXT NOT NULL, version TEXT NOT NULL, decision TEXT, approved TEXT,
    PRIMARY KEY(file_id, row_no)
);
CREATE TABLE IF NOT EXISTS journal (
    at TEXT NOT NULL, file_id TEXT, row_no INTEGER, stage TEXT NOT NULL,
    status TEXT NOT NULL, code TEXT
);
"""


def retry_files(operation, *args, **kwargs):
    for attempt in range(3):
        try:
            return operation(*args, **kwargs)
        except OSError as error:
            if attempt == 2 or getattr(error, "winerror", None) not in NETWORK_ERRORS:
                raise
            sleep(30)


@contextmanager
def run_lock(path):
    with path.open("a+b") as handle:
        try:
            if os.name == "nt":
                from msvcrt import LK_NBLCK, locking
                handle.seek(0)
                locking(handle.fileno(), LK_NBLCK, 1)
            else:
                from fcntl import LOCK_EX, LOCK_NB, flock
                flock(handle, LOCK_EX | LOCK_NB)
        except OSError:
            yield False
        else:
            yield True


def journal(db, stage, status, code=None, file_id=None, row_no=None):
    db.execute("INSERT INTO journal VALUES (?, ?, ?, ?, ?, ?)",
               (datetime.now(IRKUTSK).isoformat(), file_id, row_no, stage, status, code))


def to_json(value):
    return json.dumps(value, ensure_ascii=False)


def snapshot(path):
    with path.open("rb") as handle:
        data = handle.read(MAX_BYTES + 1)
        digest = sha256(data)
        if len(data) > MAX_BYTES:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
            data = None
    return digest.hexdigest(), data


def archive(path, directory):
    retry_files(directory.mkdir, parents=True, exist_ok=True)
    target = directory / path.name
    suffix = 1
    while retry_files(target.exists):
        target = directory / f"{path.stem}.{suffix}{path.suffix}"
        suffix += 1
    retry_files(path.rename, target)


def notify(config, db, counts):
    if not counts:
        return
    message = "\n".join(f"{key} – {value}" for key, value in counts.items())
    print(message)
    if not all(config.get(key) for key in ("host", "sender", "recipient")):
        with db:
            journal(db, "Уведомление", ERROR, "MAIL_NOT_CONFIGURED")
        return
    mail = EmailMessage()
    mail["From"] = config["sender"]
    mail["To"] = config["recipient"]
    mail["Subject"] = "ДомТорг – результат цикла категоризации"
    mail.set_content(message)
    try:
        password = read_windows_credential(config["sender"])
        if not password:
            raise ValueError
        context = ssl.create_default_context()
        for attempt in range(3):
            try:
                with smtplib.SMTP_SSL(config["host"], config.get("port", 465),
                                      timeout=30, context=context) as server:
                    server.login(config["sender"], password)
                    server.send_message(mail)
                with db:
                    journal(db, "Уведомление", "Отправлено")
                return
            except (smtplib.SMTPAuthenticationError, ssl.SSLCertVerificationError):
                break
            except OSError:
                if attempt < 2:
                    sleep(30)
    except (ValueError, ImportError):
        pass
    with db:
        journal(db, "Уведомление", ERROR, "MAIL_FAILED")


def review_rows(path):
    book = load_workbook(BytesIO(retry_files(path.read_bytes)), read_only=True,
                         data_only=False, keep_links=False)
    try:
        file_id = book["Сведения"]["B1"].value
        values = book["Категоризация"].values
        if not isinstance(file_id, str) or next(values, ())[:len(REPORT)] != REPORT:
            raise ValueError("Проверьте структуру отчёта.")
        rows = {}
        for row in values:
            if not any(row):
                continue
            if type(row[0]) is not int or row[0] in rows:
                raise ValueError("Проверьте номера строк отчёта.")
            rows[row[0]] = row[19:]
        return file_id, rows
    finally:
        book.close()


def employee_choice(ref, data, fields):
    decision, *edits = (*fields, *([None] * 7))[:7]
    decision = clean(decision)
    if decision not in DECISIONS:
        raise ValueError("Выберите решение из списка.")
    if decision == "Отклонить":
        return decision, None
    sku_source = data["pos"][0]
    if not sku_source:
        raise ValueError("Запросите у поставщика строку с артикулом.")
    result = data["result"]
    attrs = data["attrs"]
    if decision == "Исправить":
        category, sku, name, unit, coef, corrected_attrs = map(clean, edits)
        item = ref.items.get(sku)
        if item:
            name = name or item.name
            category = category or item.category
            unit = unit or item.unit
        updates = attributes(corrected_attrs)
        if len(updates) != len(parts(corrected_attrs)):
            raise ValueError("Запишите характеристики в формате ключ=значение через точку с запятой.")
        attrs = {**attrs, **updates}
    else:
        if result["result"] == CHECK:
            raise ValueError("Для этой строки укажите исправленные значения.")
        category, sku, name, unit, coef = (result[k] for k in ("category", "sku", "name", "unit", "coef"))
        if not sku and result["result"] != NEW:
            raise ValueError("Проверьте артикул или признаки новой позиции.")
    validate_result(ref, attrs, category, sku, name, unit, coef)
    return decision, dict(category=category, sku=sku, name=name, unit=unit, coef=factor(coef),
                          attrs={**(ref.items[sku].attrs if sku else {}), **attrs})


class Robot:
    def __init__(self, config_path, db, model=None):
        self.config = tomllib.loads(config_path.read_text(encoding="utf-8"))
        base = config_path.parent
        paths = self.config["paths"]
        exchange = base / paths["exchange"]
        self.inbox, self.reports, self.reviewed, self.uploads, self.archive = (
            exchange / name for name in ("Входящие файлы", "Отчёты", "Проверенные", "Загрузка", "Архив"))
        self.reference_path = base / paths["references"]
        g = self.config["gigachat"]
        ca = g.get("ca_bundle_file")
        self.model = model if model is not None else GigaChatClient(GigaChatConfig(
            g["scope"], g["model"], g["base_url"], g["verify_ssl_certs"], base / ca if ca else None))
        self.threshold = number(self.config["robot"]["threshold"])
        if not 0 <= self.threshold <= 1:
            raise ValueError("Уверенность должна быть от 0 до 1.")
        self.db = db
        self.counts = Counter()

    def run(self):
        try:
            for directory in (self.inbox, self.reports, self.reviewed, self.uploads, self.archive):
                retry_files(directory.mkdir, parents=True, exist_ok=True)
            content = retry_files(self.reference_path.read_bytes)
            self.version = sha256(content).hexdigest()
            self.ref = References(load_workbook(BytesIO(content), keep_links=False))
        except (OSError, BadZipFile, InvalidFileException, ValueError, KeyError) as error:
            print(f"Проверьте книгу справочников и папки обмена. {error}", file=sys.stderr)
            with self.db:
                journal(self.db, "Загрузка справочников", ERROR, "REFERENCES_INVALID")
            self.counts["Ошибок справочников или папок"] += 1
            notify(self.config.get("mail", {}), self.db, self.counts)
            return 1

        with self.db:
            journal(self.db, "Загрузка справочников", "Загружены", self.version)
        try:
            self.process_reviews()
            files = [(self.arrival(path), path) for folder in retry_files(list_children, self.inbox)
                     if retry_files(folder.is_dir) for path in retry_files(list_children, folder)
                     if retry_files(path.is_file) and not path.name.startswith(("~$", "."))]
            for _, path in sorted(files):
                try:
                    self.process_file(path)
                except OSError:
                    with self.db:
                        journal(self.db, "Чтение файла", ERROR, "INPUT_ACCESS")
                    self.counts["Ошибок доступа к файлу"] += 1
        except OSError:
            with self.db:
                journal(self.db, "Обмен", ERROR, "EXCHANGE_ACCESS")
            self.counts["Ошибок доступа к папкам обмена"] += 1
            return_code = 1
        else:
            return_code = 0
        finally:
            self.ref.book.close()
            with self.db:
                self.db.execute("DELETE FROM journal WHERE at < ?",
                                ((datetime.now(IRKUTSK) - timedelta(days=30)).isoformat(),))
            notify(self.config.get("mail", {}), self.db, self.counts)
        return return_code

    @staticmethod
    def arrival(path):
        stat = retry_files(path.stat)
        return stat.st_birthtime_ns if hasattr(stat, "st_birthtime_ns") else stat.st_mtime_ns

    def set_status(self, file_id, status, stage, code=None):
        with self.db:
            self.db.execute("UPDATE files SET status=? WHERE id=?", (status, file_id))
            journal(self.db, stage, status, code, file_id)

    def archive_input(self, path, file_id, duplicate=False):
        destination = self.archive / path.parent.name / ("Повторы" if duplicate else "")
        try:
            if retry_files(snapshot, path)[0] != file_id:
                raise OSError("Файл изменился во время обработки.")
            archive(path, destination)
        except OSError:
            with self.db:
                journal(self.db, "Архив", ERROR, "ARCHIVE_INPUT", file_id)
            self.counts["Ошибок архивации"] += 1
        else:
            if not duplicate:
                with self.db:
                    self.db.execute("UPDATE files SET archived=1 WHERE id=?", (file_id,))

    def process_file(self, path):
        file_id, content = retry_files(snapshot, path)
        supplier = path.parent.name
        stored = self.db.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()
        if stored and stored["status"] in {REPORTED, PARTIAL, REVIEWED}:
            self.archive_input(path, file_id, bool(stored["archived"]) or stored["supplier"] != supplier)
            return
        if stored and stored["status"] == ERROR and stored["archived"]:
            self.archive_input(path, file_id, True)
            return
        if stored and stored["supplier"] != supplier and retry_files(
                (self.inbox / stored["supplier"] / stored["name"]).exists):
            self.archive_input(path, file_id, True)
            return
        if (stored and stored["status"] == SETUP and stored["version"] == self.version
                and stored["supplier"] == supplier and Path(stored["name"]).suffix == path.suffix):
            self.counts["Файлов требуют настройки"] += 1
            return
        if not stored:
            report_name = f"{supplier[:23]} – {path.stem[:22]} – {file_id}.xlsx"
            with self.db:
                self.db.execute("INSERT INTO files (id,supplier,name,status,version,report) VALUES (?,?,?,?,?,?)",
                                (file_id, supplier, path.name, REGISTERED, self.version, report_name))
                journal(self.db, "Регистрация", REGISTERED, file_id=file_id)
        else:
            report_name = stored["report"]
            with self.db:
                self.db.execute("UPDATE files SET supplier=?,name=?,version=? WHERE id=?",
                                (supplier, path.name, self.version, file_id))
                if stored["supplier"] != supplier:
                    self.db.execute("DELETE FROM positions WHERE file_id=?", (file_id,))

        fmt = path.suffix.lower().lstrip(".")
        profile = self.ref.profiles.get(supplier)
        try:
            if fmt not in FORMATS or profile is None or profile[0] != fmt:
                raise ProfileError("Формат или поставщик не совпадает с профилем.")
            if content is None:
                raise ProfileError("Файл превышает 10 МБ.")
            positions = read_positions(content, profile)
        except ProfileError:
            self.set_status(file_id, SETUP, "Чтение файла", "INPUT_PROFILE")
            self.counts["Файлов требуют настройки"] += 1
            return
        except InputError:
            self.set_status(file_id, ERROR, "Чтение файла", "INPUT_DAMAGED")
            self.counts["Повреждённых файлов"] += 1
            self.archive_input(path, file_id)
            return

        saved = self.positions(file_id)
        unused = saved.keys() - {row_no for row_no, _, _ in positions}
        with self.db:
            self.db.executemany("DELETE FROM positions WHERE file_id=? AND row_no=?",
                                ((file_id, row_no) for row_no in unused))
        for row_no in unused:
            del saved[row_no]
        for row_no, raw, pos in positions:
            if row_no in saved and saved[row_no]["pos"] == list(pos):
                continue
            result, attrs, unit = categorize(self.ref, supplier, pos, self.model, self.threshold)
            data = dict(result=result, raw=raw, pos=pos, attrs=attrs, unit=unit)
            with self.db:
                self.db.execute("INSERT OR REPLACE INTO positions (file_id,row_no,data,version) VALUES (?,?,?,?)",
                                (file_id, row_no, to_json(data), self.version))
                journal(self.db, "Категоризация", result["result"], result["issue"] or None, file_id, row_no)
            saved[row_no] = data

        counts = Counter(row["result"]["result"] for row in saved.values())
        partial = any(row["result"]["issue"] == "LLM_UNAVAILABLE" for row in saved.values())
        try:
            self.write_report(file_id, supplier, report_name, sorted(saved.items()))
        except OSError:
            self.set_status(file_id, ERROR, "Отчёт", "REPORT_WRITE")
            self.counts["Ошибок записи отчёта"] += 1
            return
        status = PARTIAL if partial else REPORTED
        with self.db:
            self.db.execute("UPDATE files SET status=?,counts=? WHERE id=?", (status, to_json(counts), file_id))
            journal(self.db, "Отчёт", status, file_id=file_id)
        self.counts["Новых отчётов"] += 1
        self.counts["Позиций требуют решения"] += sum(counts.values())
        if partial:
            self.counts["Частично обработанных файлов"] += 1
        self.archive_input(path, file_id)

    def positions(self, file_id, only_pending=False):
        query = "SELECT row_no,data FROM positions WHERE file_id=?"
        if only_pending:
            query += " AND decision IS NULL"
        return {row["row_no"]: json.loads(row["data"]) for row in self.db.execute(query, (file_id,))}

    def write_report(self, file_id, supplier, name, positions, comments=None):
        source = self.db.execute("SELECT name FROM files WHERE id=?", (file_id,)).fetchone()[0]
        book = report_book(file_id, supplier, self.version, self.ref.leaves, positions, comments, source)
        retry_files(save_book, book, self.reports / name)

    def process_reviews(self):
        for path in sorted(retry_files(list_children, self.reviewed)):
            if path.suffix.lower() != ".xlsx" or path.name.startswith(("~$", ".")):
                continue
            try:
                self.process_review(path)
            except (InvalidFileException, KeyError, ValueError, BadZipFile):
                with self.db:
                    journal(self.db, "Решения сотрудника", ERROR, "REVIEW_FORMAT")
                self.counts["Некорректных проверенных отчётов"] += 1
            except OSError:
                with self.db:
                    journal(self.db, "Решения сотрудника", ERROR, "REVIEW_WRITE")
                self.counts["Отложенных проверенных отчётов"] += 1
                # При невозможности сохранить книгу следующая попытка начнётся с чистой копии.
                break

    def process_review(self, path):
        file_id, rows = review_rows(path)
        file = self.db.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()
        if not file or file["status"] not in {REPORTED, PARTIAL, REVIEWED}:
            raise ValueError("Отчёт не зарегистрирован.")
        pending = self.positions(file_id, only_pending=True)
        if rows.keys() - self.positions(file_id).keys():
            raise ValueError("В отчёте есть неизвестные строки.")
        decisions, comments, links = {}, {}, {}
        supplier = file["supplier"]
        for row_no, data in sorted(pending.items()):
            if not rows.get(row_no) or not rows[row_no][0]:
                continue
            try:
                decision, result = employee_choice(self.ref, data, rows[row_no])
                if result and result["sku"]:
                    key = (supplier, data["pos"][0])
                    link = (result["sku"], result["coef"])
                    if key in links and links[key] != link:
                        raise ValueError("Для артикула уже указано другое соответствие в этом отчёте.")
                    links[key] = link
            except ValueError as error:
                comments[row_no] = f"Решение не учтено. {error}"
            else:
                decisions[row_no] = (decision, result)

        if links:
            retry_files(save_links, self.ref, self.reference_path, links)
            self.version = retry_files(snapshot, self.reference_path)[0]
        with self.db:
            for row_no, (decision, result) in decisions.items():
                self.db.execute("UPDATE positions SET decision=?,approved=? WHERE file_id=? AND row_no=?",
                                (decision, to_json(result) if result else None, file_id, row_no))
                journal(self.db, "Решения сотрудника", decision, file_id=file_id, row_no=row_no)
            for row_no in comments:
                journal(self.db, "Решения сотрудника", CHECK, "REVIEW_INVALID", file_id, row_no)

        approved = [(json.loads(data)["pos"][0], json.loads(result)) for data, result in self.db.execute(
            "SELECT data,approved FROM positions WHERE file_id=? AND approved IS NOT NULL ORDER BY row_no",
            (file_id,))]
        retry_files(save_book, upload_book(supplier, approved), self.uploads / file["report"])
        if remaining := {k: v for k, v in pending.items() if k not in decisions}:
            self.write_report(file_id, supplier, file["report"], sorted(remaining.items()), comments)
            self.counts["Позиций требуют решения"] += len(remaining)
        else:
            self.set_status(file_id, REVIEWED, "Решения сотрудника")
        archive(path, self.archive / "Проверенные" / file_id[:16])
        self.counts["Учтено решений сотрудника"] += len(decisions)


def list_children(path):
    return list(path.iterdir())


def run(config_path, *, scheduled=False, model=None):
    now = datetime.now(IRKUTSK).time()
    if scheduled and not time(9) <= now < time(18, 15):
        return 0
    try:
        config = tomllib.loads(config_path.read_text(encoding="utf-8"))
        state = config_path.parent / config["paths"]["state"]
        state.parent.mkdir(parents=True, exist_ok=True)
        with run_lock(state.with_suffix(".lock")) as locked:
            if not locked:
                print("Другой экземпляр уже выполняет цикл.")
                return 0
            with closing(sqlite3.connect(state)) as db:
                db.row_factory = sqlite3.Row
                db.executescript("PRAGMA foreign_keys=ON;\n" + DATABASE)
                return Robot(config_path, db, model).run()
    except (OSError, ValueError, TypeError, KeyError, sqlite3.Error) as error:
        print(f"Не удалось запустить цикл. {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Категоризация товаров ДомТорг")
    parser.add_argument("--config", type=Path, default=Path(__file__).with_name("config.toml"))
    parser.add_argument("--scheduled", action="store_true")
    parser.add_argument("--interval", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.interval:
        print(tomllib.loads(args.config.read_text(encoding="utf-8"))["robot"]["interval_minutes"])
        sys.exit(0)
    sys.exit(run(args.config.resolve(), scheduled=args.scheduled))
