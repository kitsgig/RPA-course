import argparse
import csv
from pathlib import Path
import shutil
import tomllib
from xml.etree import ElementTree

from catalog import SCHEMA, SOURCE
from workbooks import add_sheet, report_book, save_book, upload_book, workbook


PROFILES = (
    ("ЭлектроСнаб", "xlsx", "Прайс", 2, "", "", "", *SOURCE),
    ("МастерДом", "csv", "", 1, "cp1251", ";", "", "Код", "Товар", "Ед.",
     "Марка", "Параметры", "EAN"),
    ("БытОпт", "xml", "", "", "", "", "./items/item", "sku", "name", "unit",
     "brand", "specs", "barcode"),
)

TYPES = (
    ("LIGHT.LED", "Освещение", "Лампы", "Светодиодные лампы", "мощность;цоколь",
     "лампа led;лампа светодиодная", "гирлянда"),
    ("TOOL.DRIVER", "Инструменты", "Ручной инструмент", "Отвёртки", "предметы",
     "набор отвёрток", ""),
    ("ELEC.EXT", "Электротовары", "Удлинители", "Бытовые удлинители", "розетки;длина",
     "удлинитель", ""),
    ("HOME.BOTTLE", "Товары для дома", "Посуда", "Бутылки для воды", "объём",
     "бутылка для воды", ""),
    ("HOME.TOWEL", "Товары для дома", "Текстиль", "Полотенца", "материал",
     "полотенце", "бумажное"),
    ("HOME.BOX", "Товары для дома", "Хранение", "Контейнеры", "объём",
     "контейнер", ""),
)

ITEMS = (
    ("Д-01025", "Лампа светодиодная Navigator E27 10 Вт 4000 К", "LIGHT.LED", "шт.",
     "460000000001", "бренд=Navigator;цоколь=E27;мощность=10 Вт;цветовая температура=4000 К"),
    ("Д-03112", "Набор отвёрток 6 предметов", "TOOL.DRIVER", "компл.",
     "460000000002", "предметы=6"),
    ("Д-04001", "Удлинитель 3 розетки 3 м", "ELEC.EXT", "шт.",
     "460000000003", "розетки=3;длина=3 м"),
    ("Д-05001", "Бутылка для воды 500 мл", "HOME.BOTTLE", "шт.",
     "460000000004", "объём=500 мл"),
    ("Д-06001", "Полотенце хлопковое", "HOME.TOWEL", "шт.",
     "460000000005", "материал=хлопок"),
    ("Д-07001", "Контейнер 2 л", "HOME.BOX", "шт.", "460000000006", "объём=2 л"),
)

EXAMPLES = (
    (("ЭС-4410", "Лампа LED Navigator 10W E27 4000K, упак. 10 шт", "уп.", "Navigator",
      "", ""), "LIGHT.LED", "Д-01025"),
    (("ЭС-7801", "Набор отвёрток 6 предм.", "набор", "", "", ""), "TOOL.DRIVER", "Д-03112"),
    (("ЭС-5120", "Удлинитель 3 роз. 5м с заземл.", "шт", "", "", ""), "ELEC.EXT", ""),
    (("ЭС-9003", "Товар акционный 2шт", "шт", "", "", ""), "", ""),
)


def references():
    result = workbook()
    categories, rules = {}, []
    for code, group, subgroup, name, required, include, exclude in TYPES:
        root = code.split(".")[0]
        mid = code + ".GROUP"
        categories[root] = (root, group, "", "")
        categories[mid] = (mid, subgroup, root, "")
        categories[code] = (code, name, mid, required)
        rules.append((code, include, exclude, 10))
    rows = {
        "Категории": categories.values(), "Правила": rules, "Каталог": ITEMS,
        "Единицы": (("шт.", "шт;штука;pcs"), ("упак.", "уп.;уп;упаковка"),
                    ("компл.", "комплект;набор"), ("кг", "кг.;kg"),
                    ("л", "л.;l"), ("м", "м.;m")),
        "Соответствия": (("ЭлектроСнаб", "ЭС-4410", "Д-01025", 10),),
        "Поставщики": PROFILES,
    }
    for title, header in SCHEMA.items():
        add_sheet(result, title, header, rows[title])
    return result


def examples(index, count, model_rows):
    sample = list(EXAMPLES) if index == 0 else []
    for i in range(len(sample), count):
        sku, title, category, unit, barcode, attrs = ITEMS[(i + index * 2) % len(ITEMS)]
        if i < model_rows:
            barcode = ""
        row = (f"{index + 1}-{i + 1:04}", title, unit, "", attrs, barcode)
        sample.append((row, category, sku))
    return sample


def write_supplier(path, index, sample):
    _, fmt, sheet, _, encoding, delimiter, _, *fields = PROFILES[index]
    rows = [row for row, _, _ in sample]
    if fmt == "xlsx":
        book = workbook()
        page = add_sheet(book, sheet, fields, rows)
        page.insert_rows(1)
        page["A1"] = "Учебный перечень ЭлектроСнаб"
        page.row_dimensions[2].height = 32
        page.freeze_panes = "B3"
        page.auto_filter.ref = f"A2:F{page.max_row}"
        save_book(book, path)
    elif fmt == "csv":
        with path.open("w", newline="", encoding=encoding) as stream:
            writer = csv.writer(stream, delimiter=delimiter)
            writer.writerow(fields)
            writer.writerows(rows)
    else:
        root = ElementTree.Element("supplier")
        items = ElementTree.SubElement(root, "items")
        for row in rows:
            item = ElementTree.SubElement(items, "item")
            for field, value in zip(fields, row, strict=True):
                ElementTree.SubElement(item, field).text = value
        ElementTree.ElementTree(root).write(path, encoding="utf-8", xml_declaration=True)


def generate(destination):
    destination.mkdir(parents=True, exist_ok=True)
    reference_path = destination / "Справочники каталога.xlsx"
    if reference_path.exists():
        raise ValueError("Для учебных данных укажите новую папку.")
    save_book(references(), reference_path)
    save_book(report_book("", "", "", {}), destination / "Шаблон отчёта.xlsx")
    save_book(upload_book(""), destination / "Шаблон загрузки.xlsx")
    control = []
    for index, ((supplier, fmt, *_), count) in enumerate(zip(PROFILES, (68, 66, 66), strict=True)):
        folder = destination / "Поставщики" / supplier
        folder.mkdir(parents=True, exist_ok=True)
        sample = examples(index, count, 8)
        write_supplier(folder / f"перечень.{fmt}", index, sample)
        control.extend((supplier, row[0], category, sku) for row, category, sku in sample)
    with (destination / "Эталон 200.csv").open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(("Поставщик", "Артикул поставщика", "Код категории", "Внутренний артикул"))
        writer.writerows(control)


def prepare(config_path, performance=False):
    paths = tomllib.loads(config_path.read_text(encoding="utf-8"))["paths"]
    exchange = config_path.parent / paths["exchange"]
    ref = config_path.parent / paths["references"]
    inbox = exchange / "Входящие файлы"
    if ref.exists() or inbox.exists():
        raise ValueError("Укажите конфигурацию для чистой папки обмена. Учебный запуск не меняет имеющиеся файлы.")
    samples = Path(__file__).with_name("samples")
    ref.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(samples / "Справочники каталога.xlsx", ref)
    if performance:
        folder = inbox / PROFILES[0][0]
        folder.mkdir(parents=True)
        write_supplier(folder / "нагрузка.xlsx", 0, examples(0, 500, 101))
    else:
        shutil.copytree(samples / "Поставщики", inbox)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Подготовка учебных данных ДомТорг")
    parser.add_argument("--config", type=Path, default=Path(__file__).with_name("config.toml"))
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--generate", type=Path)
    mode.add_argument("--performance", action="store_true")
    args = parser.parse_args()
    try:
        if args.generate:
            generate(args.generate)
        else:
            prepare(args.config.resolve(), args.performance)
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, f"{error}\n")
