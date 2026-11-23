import argparse
from collections import Counter
from contextlib import closing, redirect_stdout
import csv
from io import StringIO
import json
from pathlib import Path
import shutil
import sqlite3
from tempfile import TemporaryDirectory
from time import perf_counter
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from catalog import CHECK, MATCHED, PROPOSED, InputError, References, read_positions
from demo import ITEMS, prepare
from gigachat import GigaChatError, GigaChatClient, GigaChatConfig, GigaChatTimeout
from robot import PARTIAL, REVIEWED, SETUP, run, run_lock


ROOT = Path(__file__).parent
EXCHANGE = "Обмен"


class ModelStub:
    def __init__(self):
        self.calls = 0

    def complete(self, request):
        self.calls += 1
        data = json.loads(request["messages"][1]["content"])
        name = data["name"]
        assert not {"sku", "supplier", "barcode", "price"} & data.keys()
        assert len(data["candidates"]) <= 20 and "functions" not in request
        coef, confidence = 1, 0.95
        if "заземл." in name:
            category, sku, title, unit = "ELEC.EXT", "", "Удлинитель 3 розетки 5 м с заземлением", "шт."
        elif "акционный" in name:
            sku, title, category, unit, *_ = ITEMS[0]
            confidence = 0.2
        elif "6 предм." in name:
            sku, title, category, unit, *_ = ITEMS[1]
        else:
            sku, title, category, unit, *_ = next(item for item in ITEMS if item[1] == name)
        return json.dumps(dict(category=category, sku=sku, name=title, unit=unit, new=not sku,
                               coef=coef, confidence=confidence))


def new_config(folder, performance=False):
    config = folder / "config.toml"
    shutil.copy2(ROOT / "config.toml", config)
    shutil.copytree(ROOT / "certs", folder / "certs")
    prepare(config, performance)
    return config


def read_stats(folder):
    with closing(sqlite3.connect(folder / ".runtime/state.sqlite3")) as db:
        rows = db.execute("SELECT supplier,data FROM positions JOIN files ON file_id=id")
        return [(supplier, json.loads(data)) for supplier, data in rows]


def metrics(rows):
    with (ROOT / "samples/Эталон 200.csv").open(encoding="utf-8-sig", newline="") as stream:
        expected = {(supplier, sku): category for supplier, sku, category, _
                    in list(csv.reader(stream, delimiter=";"))[1:]}
    scored = checked = correct = 0
    for supplier, row in rows:
        result = row["result"]
        if result["result"] in {MATCHED, PROPOSED}:
            scored += 1
            correct += result["category"] == expected[(supplier, row["pos"][0])]
        elif result["result"] == CHECK:
            checked += 1
    return correct, scored, checked


class Acceptance(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.config = new_config(self.folder)
        self.exchange = self.folder / EXCHANGE
        self.model = ModelStub()

    def cycle(self, model=None):
        with redirect_stdout(StringIO()):
            self.assertEqual(run(self.config, model=self.model if model is None else model), 0)

    def db(self, query, args=()):
        with closing(sqlite3.connect(self.folder / ".runtime/state.sqlite3")) as db:
            return db.execute(query, args).fetchall()

    def test_200_positions_reports_review_and_deduplication(self):
        self.cycle()
        rows = read_stats(self.folder)
        self.assertEqual(len(rows), 200)
        self.assertEqual(metrics(rows), (198, 198, 1))
        self.assertEqual(self.model.calls, 23)
        report_name, = self.db("SELECT report FROM files WHERE supplier='ЭлектроСнаб'")[0]
        report = self.exchange / "Отчёты" / report_name
        with closing(load_workbook(report)) as book:
            sheet = book["Категоризация"]
            self.assertEqual(sheet["P2"].value, 10)
            sheet["T2"] = "Принять"
            sheet["T3"] = "Исправить"
            sheet["V3"] = "Д-03112"
            sheet["Y3"] = 1
            sheet["T4"] = "Принять"
            sheet["T5"] = "Отклонить"
            sheet["T6"] = "Исправить"
            sheet["V6"] = "несуществующий артикул"
            sheet["Y6"] = -1
            book.save(self.exchange / "Проверенные" / report_name)
        self.cycle()
        with closing(load_workbook(self.exchange / "Загрузка" / report_name)) as book:
            self.assertEqual(book["Загрузка"].max_row, 3)
            self.assertEqual(book["Новые товары"].max_row, 2)
        with closing(load_workbook(self.exchange / "Справочники каталога.xlsx")) as book:
            self.assertIn(("ЭлектроСнаб", "ЭС-7801", "Д-03112", 1), list(book["Соответствия"].values))
        with closing(load_workbook(report)) as book:
            sheet = book["Категоризация"]
            self.assertIn("Решение не учтено", sheet["S2"].value)
            for cells in sheet.iter_rows(min_row=2):
                cells[19].value = "Отклонить"
            book.save(self.exchange / "Проверенные" / report_name)
        self.cycle()
        self.assertEqual(self.db("SELECT status FROM files WHERE report=?", (report_name,)), [(REVIEWED,)])
        archived = self.exchange / "Архив/ЭлектроСнаб/перечень.xlsx"
        shutil.copy2(archived, self.exchange / "Входящие файлы/ЭлектроСнаб/копия.xlsx")
        self.cycle()
        self.assertEqual(len(self.db("SELECT id FROM files")), 3)
        self.assertEqual(self.model.calls, 23)
        self.assertTrue((archived.parent / "Повторы/копия.xlsx").is_file())

    def test_resume_after_report_failure(self):
        from robot import save_book

        def failing_save(book, path):
            if path.parent.name == "Отчёты":
                raise PermissionError
            save_book(book, path)

        with patch("robot.save_book", side_effect=failing_save):
            self.cycle()
        self.assertEqual(self.db("SELECT DISTINCT status FROM files"), [("Ошибка",)])
        self.assertEqual(len(read_stats(self.folder)), 200)
        self.cycle()
        self.assertEqual(len(list((self.exchange / "Отчёты").glob("*.xlsx"))), 3)
        self.assertEqual(self.model.calls, 23)

    def test_partial_profile_and_xml_security(self):
        class Unavailable:
            def complete(self, request):
                raise GigaChatError("Учебный отказ модели.")

        unknown = self.exchange / "Входящие файлы/Другой"
        unknown.mkdir()
        (unknown / "вход.csv").write_bytes(b"Code;Name\nA001;Lamp")
        self.cycle(Unavailable())
        self.assertEqual(self.db("SELECT COUNT(*) FROM files WHERE status=?", (PARTIAL,)), [(3,)])
        self.assertEqual(self.db("SELECT COUNT(*) FROM files WHERE status=?", (SETUP,)), [(1,)])
        self.assertTrue((unknown / "вход.csv").is_file())
        ref_path = self.exchange / "Справочники каталога.xlsx"
        with closing(load_workbook(ref_path)) as book:
            profile = References(book).profiles["БытОпт"]
            with self.assertRaises(InputError):
                read_positions(b'<!DOCTYPE supplier [<!ENTITY x SYSTEM "file:///secret">]>'
                               b'<supplier>&x;</supplier>', profile)
            book["Поставщики"].append(("Другой", "csv", None, 1, "utf-8", ";", None, "Code", "Name"))
            book.save(ref_path)
        self.cycle(Unavailable())
        self.assertFalse((unknown / "вход.csv").exists())
        self.assertEqual(self.db("SELECT status FROM files WHERE supplier='Другой'"), [(PARTIAL,)])

    def test_resume_at_first_unprocessed_position(self):
        from robot import categorize

        count = 0

        def interrupt_after_three(*args):
            nonlocal count
            count += 1
            if count == 4:
                raise KeyboardInterrupt
            return categorize(*args)

        with patch("robot.categorize", side_effect=interrupt_after_three):
            with self.assertRaises(KeyboardInterrupt):
                self.cycle()
        self.assertEqual(len(read_stats(self.folder)), 3)
        self.cycle()
        self.assertEqual(len(read_stats(self.folder)), 200)
        self.assertEqual(self.model.calls, 23)

    def test_invalid_response_retry_formulas_and_lock(self):
        with run_lock(self.folder / "test.lock") as outer:
            self.assertTrue(outer)
            with run_lock(self.folder / "test.lock") as inner:
                self.assertFalse(inner)

        class Invalid:
            calls = 0

            def complete(self, request):
                self.calls += 1
                return '{"category":"fake"}'

        model = Invalid()
        path = self.exchange / "Входящие файлы/МастерДом/формула.csv"
        path.write_text("Код;Товар;Ед.;Марка;Параметры;EAN\n=2+2;@Lamp;шт;;;\n",
                        encoding="cp1251")
        self.cycle(model)
        self.assertEqual(model.calls, 48)
        report, = self.db("SELECT report FROM files WHERE name='формула.csv'")[0]
        with closing(load_workbook(self.exchange / "Отчёты" / report)) as book:
            sheet = book["Категоризация"]
            self.assertEqual((sheet["B2"].value, sheet["B2"].data_type), ("=2+2", "s"))
            self.assertEqual(sheet["C2"].data_type, "s")
            self.assertEqual(sheet["J2"].value, CHECK)

    def test_gigachat_timeout_and_network_retry(self):
        client = GigaChatClient(GigaChatConfig("GIGACHAT_API_PERS", "GigaChat-2-Pro",
                                               "https://api.giga.chat/v1", True))
        with patch.object(client, "_opener") as opener, patch("gigachat.sleep") as wait:
            opener.open.side_effect = TimeoutError
            with self.assertRaises(GigaChatTimeout):
                client._request_json("https://api.giga.chat/v1/chat/completions", b"{}", {})
            self.assertEqual(opener.open.call_args.kwargs["timeout"], 60)
            self.assertEqual(opener.open.call_count, 1)
            opener.open.reset_mock(side_effect=True)
            opener.open.side_effect = OSError
            with self.assertRaises(GigaChatError):
                client._request_json("https://api.giga.chat/v1/chat/completions", b"{}", {})
            self.assertEqual(opener.open.call_count, 3)
            self.assertEqual([call.args for call in wait.call_args_list], [(30,), (30,)])


def benchmark(live, performance, destination=None):
    model = None if live else ModelStub()
    started = perf_counter()
    if destination:
        destination.mkdir(parents=True, exist_ok=True)
        config = new_config(destination, performance)
        code = run(config, model=model)
        rows = read_stats(destination)
    else:
        with TemporaryDirectory() as folder:
            config = new_config(Path(folder), performance)
            code = run(config, model=model)
            rows = read_stats(Path(folder))
    elapsed = perf_counter() - started
    print(f"Позиций {len(rows)}, время {elapsed:.2f} с")
    issues = Counter(row["result"]["issue"] for _, row in rows if row["result"]["issue"])
    if issues:
        print("Коды отклонений – " + ", ".join(f"{code} ({count})" for code, count in sorted(issues.items())))
    if not performance:
        correct, scored, checked = metrics(rows)
        print(f"Категория верна {correct}/{scored}. Требуют проверки {checked}/{len(rows)}.")
        if len(rows) != 200 or not scored or correct / scored < 0.90 or checked / len(rows) > 0.20:
            return 1
    if performance and (len(rows) != 500 or elapsed > 30 * 60):
        return 1
    return 1 if issues["LLM_UNAVAILABLE"] else code


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--live", action="store_true")
    parser.add_argument("--performance", action="store_true")
    parser.add_argument("--output", type=Path, help="Новая папка для сохранения результатов")
    options = parser.parse_args()
    if options.live or options.performance or options.output:
        raise SystemExit(benchmark(options.live, options.performance, options.output))
    unittest.main(argv=[__file__])
