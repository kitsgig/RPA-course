from pathlib import Path
from math import ceil
from uuid import uuid4

from openpyxl import Workbook
from openpyxl.cell.cell import Cell, ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from catalog import clean


REPORT = (
    "Строка", "Артикул поставщика", "Название поставщика", "Единица поставщика",
    "Бренд", "Исходные характеристики", "Штрихкод", "Нормализованная единица",
    "Характеристики", "Результат", "Код категории", "Категория", "Внутренний артикул",
    "Внутреннее наименование", "Единица хранения", "Коэффициент",
    "Способ", "Уверенность", "Комментарий робота", "Решение сотрудника",
    "Испр. код категории", "Испр. артикул", "Испр. наименование",
    "Испр. единица", "Испр. коэффициент", "Испр. характеристики",
)
DECISIONS = ("Принять", "Исправить", "Отклонить")
UPLOAD = ("Поставщик", "Артикул поставщика", "Внутренний артикул", "Внутреннее наименование",
          "Код категории", "Единица хранения", "Коэффициент", "Характеристики")


def append_text(sheet, values):
    cells = []
    for value in values:
        if isinstance(value, str):
            value = ILLEGAL_CHARACTERS_RE.sub("", value)
            cell = Cell(sheet, value=value)
            cell.data_type = "s"
            cell.number_format = "@"
            cells.append(cell)
        else:
            cells.append(value)
    sheet.append(cells)


def add_sheet(book, title, headers, rows):
    sheet = book.create_sheet(title)
    append_text(sheet, headers)
    for row in rows:
        append_text(sheet, row)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="203650")
        cell.fill = PatternFill("solid", fgColor="E7EEF7")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sample = list(sheet.iter_rows(min_row=2, max_row=min(sheet.max_row, 41), values_only=True))
    widths = []
    for col, header in enumerate(headers, 1):
        values = [len(str(row[col - 1] or "")) for row in sample]
        width = min(48, max(12, len(header) + 3, max(values, default=0) + 2))
        widths.append(width)
        sheet.column_dimensions[get_column_letter(col)].width = width
    wrapped = Alignment(wrap_text=True, vertical="center")
    for row in sheet.iter_rows(min_row=2):
        lines = 1
        for cell, width in zip(row, widths, strict=True):
            if isinstance(cell.value, str) and len(cell.value) > width:
                cell.alignment = wrapped
                lines = max(lines, ceil(len(cell.value) / width))
        if lines > 1:
            sheet.row_dimensions[row[0].row].height = min(72, lines * 15)
    sheet.row_dimensions[1].height = 32
    sheet.freeze_panes = "B2"
    sheet.auto_filter.ref = sheet.dimensions
    return sheet


def workbook():
    book = Workbook()
    book.remove(book.active)
    return book


def save_book(book, path):
    path = Path(path)
    temp = path.with_name(f".{path.stem}.{uuid4().hex}.tmp")
    try:
        book.save(temp)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def attrs_text(attrs):
    return "; ".join(f"{key}={value}" for key, value in sorted(attrs.items()))


def report_book(file_id, supplier, version, leaves, positions=(), comments=None, source=""):
    book = workbook()
    comments = comments or {}
    rows = []
    for row_no, data in positions:
        result = data["result"]
        category = result["category"]
        rows.append((
            row_no, *data["raw"], data["unit"], attrs_text(data["attrs"]),
            result["result"], category, leaves.get(category, ""), result["sku"], result["name"],
            result["unit"], result["coef"], result["method"], result["confidence"],
            comments.get(row_no, result["comment"]),
        ))
    sheet = add_sheet(book, "Категоризация", REPORT, rows)
    for cell in sheet[1][19:]:
        cell.fill = PatternFill("solid", fgColor="DDECE3")
    choice = DataValidation(type="list", formula1='"' + ",".join(DECISIONS) + '"',
                            allow_blank=True, showErrorMessage=True,
                            errorTitle="Неизвестное решение", error="Выберите значение из списка.")
    sheet.add_data_validation(choice)
    choice.add(f"T2:T{max(2, len(rows) + 1)}")
    meta = book.create_sheet("Сведения")
    for values in (("Идентификатор", file_id), ("Поставщик", supplier),
                   ("Версия справочников", version), ("Исходный файл", source)):
        append_text(meta, values)
    meta.column_dimensions["A"].width = 28
    meta.column_dimensions["B"].width = 70
    return book


def upload_book(supplier, decisions=()):
    book = workbook()
    rows = {True: [], False: []}
    for source_sku, result in decisions:
        rows[bool(result["sku"])].append((
            supplier, source_sku, result["sku"], result["name"], result["category"],
            result["unit"], result["coef"], attrs_text(result["attrs"]),
        ))
    add_sheet(book, "Загрузка", UPLOAD, rows[True])
    add_sheet(book, "Новые товары", UPLOAD, rows[False])
    return book


def save_links(ref, path, links):
    sheet = ref.book["Соответствия"]
    existing = {(clean(supplier), clean(sku)): index for index, (supplier, sku, *_) in
                enumerate(sheet.values, 1) if index > 1 and supplier and sku}
    for key, (sku, coef) in links.items():
        if index := existing.get(key):
            sheet.cell(index, 3, sku).data_type = "s"
            sheet.cell(index, 4, coef)
        else:
            append_text(sheet, (*key, sku, coef))
    sheet.auto_filter.ref = sheet.dimensions
    save_book(ref.book, path)
    ref.links.update(links)
