# -*- coding: utf-8 -*-
"""
Скрипт создаёт файл Учет_Распиловки.xlsx для учёта распиловки леса.

Запуск:
    python build_excel.py

Требуется только библиотека openpyxl:
    pip install openpyxl
"""

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUTPUT_FILE = "Учет_Распиловки.xlsx"

LAST_ROW_RASPILOVKA = 10000
LAST_ROW_RASHODY = 10000
LAST_ROW_SVOD_MASTER = 100

# ---------------------------------------------------------------------------
# Общие стили
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
HEADER_FONT = Font(bold=True)
TOTAL_FONT = Font(bold=True)

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_MONEY = "#,##0.00"
FMT_VOLUME = "#,##0.000"
FMT_DATE = "DD.MM.YYYY"
FMT_PERCENT = "0.0%"


def style_header_row(ws, row, first_col, last_col):
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_total_row(ws, row, first_col, last_col):
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = TOTAL_FONT
        cell.fill = TOTAL_FILL
        cell.border = BORDER


def autofit_columns(ws, min_width=10, max_width=45):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            col_letter = cell.column_letter
            length = len(str(cell.value))
            if length > widths.get(col_letter, 0):
                widths[col_letter] = length
    for col_letter, length in widths.items():
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, length + 2))


def freeze_header(ws):
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Создание книги
# ---------------------------------------------------------------------------

wb = Workbook()
wb.remove(wb.active)

ws_instr = wb.create_sheet("Инструкция")
ws_settings = wb.create_sheet("Настройки")
ws_ref = wb.create_sheet("Справочник_Статей")
ws_raspil = wb.create_sheet("Распиловка")
ws_rashody = wb.create_sheet("Расходы")
ws_planfact = wb.create_sheet("ПланФакт_Расходы")
ws_master = wb.create_sheet("Свод_Мастер")
ws_period = wb.create_sheet("Свод_Период")

# ---------------------------------------------------------------------------
# Лист "Инструкция"
# ---------------------------------------------------------------------------

ws_instr["A1"] = "Как пользоваться таблицей"
ws_instr["A1"].font = Font(bold=True, size=14)

instructions = [
    "1. Заполните лист «Настройки».",
    "2. Проверьте нормативы на листе «Справочник_Статей» (итог должен быть 500 руб/м3).",
    "3. Копируйте данные по распиловке на лист «Распиловка» (столбцы A, C, D, E, F, M).",
    "4. Копируйте расходы на лист «Расходы» (столбцы A, C, D, E, F, G, H, I).",
    "5. Меняйте период в «Настройки!B9» — все своды пересчитаются.",
    "6. Смотрите итоги на листах «ПланФакт_Расходы», «Свод_Мастер», «Свод_Период».",
]

for i, line in enumerate(instructions, start=3):
    cell = ws_instr.cell(row=i, column=1, value=line)
    cell.alignment = Alignment(wrap_text=True)

ws_instr.column_dimensions["A"].width = 90
ws_instr.row_dimensions[1].height = 22
freeze_header(ws_instr)

# ---------------------------------------------------------------------------
# Лист "Настройки"
# ---------------------------------------------------------------------------

settings_header = ["Параметр", "Значение", "Ед."]
for col, title in enumerate(settings_header, start=1):
    ws_settings.cell(row=1, column=col, value=title)
style_header_row(ws_settings, 1, 1, 3)

settings_rows = [
    ("Ставка заказчика", 1500, "руб/м3"),
    ("Ставка бригады", 500, "руб/м3"),
    ("Валовая ставка мастера", "=B2-B3", "руб/м3"),
    ("План расходов мастера", 500, "руб/м3"),
    ("План зарплаты мастера", "=B4-B5", "руб/м3"),
    ("Резерв", 0, "руб/м3"),
    ("Валюта", "руб", ""),
    ("Период по умолчанию", "09.2025", "ММ.ГГГГ"),
]

for i, (param, value, unit) in enumerate(settings_rows, start=2):
    ws_settings.cell(row=i, column=1, value=param)
    cell_b = ws_settings.cell(row=i, column=2, value=value)
    ws_settings.cell(row=i, column=3, value=unit)
    if isinstance(value, (int, float)):
        cell_b.number_format = FMT_MONEY
    for col in range(1, 4):
        ws_settings.cell(row=i, column=col).border = BORDER

# B9 - период, хранится как текст
ws_settings["B9"].number_format = "@"
ws_settings["B9"] = "09.2025"

freeze_header(ws_settings)

# ---------------------------------------------------------------------------
# Лист "Справочник_Статей"
# ---------------------------------------------------------------------------

ref_header = ["Статья", "Тип", "План_руб/м3", "Порядок"]
for col, title in enumerate(ref_header, start=1):
    ws_ref.cell(row=1, column=col, value=title)
style_header_row(ws_ref, 1, 1, 4)

ref_rows = [
    ("Камаз соляра", "Переменная", 100, 1),
    ("Ленты", "Переменная", 120, 2),
    ("Перчатки", "Переменная", 10, 3),
    ("Грузовик соляра", "Переменная", 50, 4),
    ("Бензин на пилу", "Переменная", 10, 5),
    ("Подшипник на раму", "Постоянная", 20, 6),
    ("Сторожи Ирина", "Постоянная", 60, 7),
    ("Соляра", "Переменная", 50, 8),
    ("Заточка пил", "Переменная", 80, 9),
]

for i, (statya, tip, plan, order) in enumerate(ref_rows, start=2):
    ws_ref.cell(row=i, column=1, value=statya)
    ws_ref.cell(row=i, column=2, value=tip)
    c_cell = ws_ref.cell(row=i, column=3, value=plan)
    c_cell.number_format = FMT_MONEY
    ws_ref.cell(row=i, column=4, value=order)
    for col in range(1, 5):
        ws_ref.cell(row=i, column=col).border = BORDER

total_row = 11
ws_ref.cell(row=total_row, column=1, value="Итого")
ws_ref.cell(row=total_row, column=3, value="=SUM(C2:C10)").number_format = FMT_MONEY
style_total_row(ws_ref, total_row, 1, 4)

freeze_header(ws_ref)

# ---------------------------------------------------------------------------
# Лист "Распиловка"
# ---------------------------------------------------------------------------

raspil_header = [
    "Дата", "Период", "Объект", "Мастер", "Бригада", "Объем_м3",
    "Ставка_заказчика", "Начислено_всего", "Бригаде", "Мастеру_валовая",
    "План_расходов", "План_зарплаты", "Примечание",
]
for col, title in enumerate(raspil_header, start=1):
    ws_raspil.cell(row=1, column=col, value=title)
style_header_row(ws_raspil, 1, 1, len(raspil_header))

for r in range(2, LAST_ROW_RASPILOVKA + 1):
    ws_raspil.cell(row=r, column=1).number_format = FMT_DATE
    ws_raspil.cell(row=r, column=2, value=f'=IF(A{r}="","",TEXT(A{r},"MM.YYYY"))')
    ws_raspil.cell(row=r, column=6).number_format = FMT_VOLUME
    ws_raspil.cell(row=r, column=7, value="=Настройки!$B$2")
    ws_raspil.cell(row=r, column=7).number_format = FMT_MONEY
    ws_raspil.cell(row=r, column=8, value=f'=IF(F{r}="","",F{r}*G{r})')
    ws_raspil.cell(row=r, column=8).number_format = FMT_MONEY
    ws_raspil.cell(row=r, column=9, value=f'=IF(F{r}="","",F{r}*Настройки!$B$3)')
    ws_raspil.cell(row=r, column=9).number_format = FMT_MONEY
    ws_raspil.cell(row=r, column=10, value=f'=IF(F{r}="","",F{r}*Настройки!$B$4)')
    ws_raspil.cell(row=r, column=10).number_format = FMT_MONEY
    ws_raspil.cell(row=r, column=11, value=f'=IF(F{r}="","",F{r}*Настройки!$B$5)')
    ws_raspil.cell(row=r, column=11).number_format = FMT_MONEY
    ws_raspil.cell(row=r, column=12, value=f'=IF(F{r}="","",F{r}*Настройки!$B$6)')
    ws_raspil.cell(row=r, column=12).number_format = FMT_MONEY

# Тестовая строка (проверка расчётов)
ws_raspil["A2"] = "2025-09-01"
from datetime import date
ws_raspil["A2"] = date(2025, 9, 1)
ws_raspil["C2"] = "Объект 1"
ws_raspil["D2"] = "Мастер 1"
ws_raspil["E2"] = "Бригада 1"
ws_raspil["F2"] = 61.206
ws_raspil["M2"] = ""

freeze_header(ws_raspil)

# ---------------------------------------------------------------------------
# Лист "Расходы"
# ---------------------------------------------------------------------------

rashody_header = [
    "Дата", "Период", "Объект", "Мастер", "Бригада", "Статья", "Сумма",
    "Комментарий", "Чек",
]
for col, title in enumerate(rashody_header, start=1):
    ws_rashody.cell(row=1, column=col, value=title)
style_header_row(ws_rashody, 1, 1, len(rashody_header))

for r in range(2, LAST_ROW_RASHODY + 1):
    ws_rashody.cell(row=r, column=1).number_format = FMT_DATE
    ws_rashody.cell(row=r, column=2, value=f'=IF(A{r}="","",TEXT(A{r},"MM.YYYY"))')
    ws_rashody.cell(row=r, column=7).number_format = FMT_MONEY

# Data validation для столбца "Статья" (F)
dv = DataValidation(
    type="list",
    formula1="Справочник_Статей!$A$2:$A$10",
    allow_blank=True,
)
ws_rashody.add_data_validation(dv)
dv.add(f"F2:F{LAST_ROW_RASHODY}")

# Тестовые строки расходов
test_expenses = [
    (date(2025, 9, 5), "Объект 1", "Мастер 1", "Бригада 1", "Камаз соляра", 10600),
    (date(2025, 9, 5), "Объект 1", "Мастер 1", "Бригада 1", "Ленты", 7300),
    (date(2025, 9, 5), "Объект 1", "Мастер 1", "Бригада 1", "Грузовик соляра", 3050),
    (date(2025, 9, 5), "Объект 1", "Мастер 1", "Бригада 1", "Сторожи Ирина", 4000),
]
for i, (dat, obj, master, brigada, statya, summa) in enumerate(test_expenses, start=2):
    ws_rashody.cell(row=i, column=1, value=dat)
    ws_rashody.cell(row=i, column=3, value=obj)
    ws_rashody.cell(row=i, column=4, value=master)
    ws_rashody.cell(row=i, column=5, value=brigada)
    ws_rashody.cell(row=i, column=6, value=statya)
    ws_rashody.cell(row=i, column=7, value=summa)

freeze_header(ws_rashody)

# ---------------------------------------------------------------------------
# Лист "ПланФакт_Расходы"
# ---------------------------------------------------------------------------

ws_planfact["A1"] = "Период"
ws_planfact["B1"] = "=Настройки!$B$9"

top_block = [
    ("Объем, м3", "=SUMIF(Распиловка!$B:$B,$B$1,Распиловка!$F:$F)", FMT_VOLUME),
    ("План расходов всего", "=B2*Настройки!$B$5", FMT_MONEY),
    ("Факт расходов всего", "=SUMIF(Расходы!$B:$B,$B$1,Расходы!$G:$G)", FMT_MONEY),
    ("Отклонение", "=B4-B3", FMT_MONEY),
    ("Валовая мастера", "=B2*Настройки!$B$4", FMT_MONEY),
    ("Чистая зарплата мастера", "=B6-B4", FMT_MONEY),
    ("Руб/м3 факт расходов", '=IF(B2=0,"",B4/B2)', FMT_MONEY),
]
for i, (label, formula, fmt) in enumerate(top_block, start=2):
    ws_planfact.cell(row=i, column=1, value=label)
    cell = ws_planfact.cell(row=i, column=2, value=formula)
    cell.number_format = fmt

# Таблица по статьям
pf_header_row = 10
pf_header = ["Статья", "План руб/м3", "План сумма", "Факт сумма", "Отклонение", "%"]
for col, title in enumerate(pf_header, start=1):
    ws_planfact.cell(row=pf_header_row, column=col, value=title)
style_header_row(ws_planfact, pf_header_row, 1, len(pf_header))

for offset, r in enumerate(range(11, 20)):
    ref_row = offset + 2  # Справочник_Статей!row 2..10
    ws_planfact.cell(row=r, column=1, value=f"=Справочник_Статей!A{ref_row}")
    b_cell = ws_planfact.cell(row=r, column=2, value=f"=Справочник_Статей!C{ref_row}")
    b_cell.number_format = FMT_MONEY
    c_cell = ws_planfact.cell(row=r, column=3, value=f"=$B$2*B{r}")
    c_cell.number_format = FMT_MONEY
    d_cell = ws_planfact.cell(
        row=r, column=4,
        value=f"=SUMIFS(Расходы!$G:$G,Расходы!$B:$B,$B$1,Расходы!$F:$F,$A{r})",
    )
    d_cell.number_format = FMT_MONEY
    e_cell = ws_planfact.cell(row=r, column=5, value=f"=D{r}-C{r}")
    e_cell.number_format = FMT_MONEY
    f_cell = ws_planfact.cell(row=r, column=6, value=f'=IF(C{r}=0,"",D{r}/C{r}-1)')
    f_cell.number_format = FMT_PERCENT
    for col in range(1, 7):
        ws_planfact.cell(row=r, column=col).border = BORDER

pf_total_row = 20
ws_planfact.cell(row=pf_total_row, column=1, value="Итого")
b20 = ws_planfact.cell(row=pf_total_row, column=2, value="=SUM(B11:B19)")
b20.number_format = FMT_MONEY
c20 = ws_planfact.cell(row=pf_total_row, column=3, value="=SUM(C11:C19)")
c20.number_format = FMT_MONEY
d20 = ws_planfact.cell(row=pf_total_row, column=4, value="=SUM(D11:D19)")
d20.number_format = FMT_MONEY
e20 = ws_planfact.cell(row=pf_total_row, column=5, value="=SUM(E11:E19)")
e20.number_format = FMT_MONEY
f20 = ws_planfact.cell(row=pf_total_row, column=6, value='=IF(C20=0,"",D20/C20-1)')
f20.number_format = FMT_PERCENT
style_total_row(ws_planfact, pf_total_row, 1, 6)

for col in range(1, 7):
    ws_planfact.cell(row=1, column=col).font = HEADER_FONT

ws_planfact.column_dimensions["A"].width = 25

freeze_header(ws_planfact)

# ---------------------------------------------------------------------------
# Лист "Свод_Мастер"
# ---------------------------------------------------------------------------

ws_master["A1"] = "Период"
ws_master["B1"] = "=Настройки!$B$9"
ws_master["A1"].font = HEADER_FONT

master_header_row = 2
master_header = [
    "Мастер", "Объем", "Начислено", "Бригаде", "Мастеру валовая",
    "План расходов", "Факт расходов", "Отклонение", "Чистая зарплата",
    "Руб/м3 факт", "План зарплаты", "Отклонение зарплаты",
]
for col, title in enumerate(master_header, start=1):
    ws_master.cell(row=master_header_row, column=col, value=title)
style_header_row(ws_master, master_header_row, 1, len(master_header))

money_cols_master = [3, 4, 5, 6, 7, 8, 9, 11, 12]

for r in range(3, LAST_ROW_SVOD_MASTER + 1):
    ws_master.cell(row=r, column=2, value=(
        f'=SUMIFS(Распиловка!$F:$F,Распиловка!$B:$B,$B$1,Распиловка!$D:$D,$A{r})'
    ))
    ws_master.cell(row=r, column=2).number_format = FMT_VOLUME
    ws_master.cell(row=r, column=3, value=f"=B{r}*Настройки!$B$2")
    ws_master.cell(row=r, column=4, value=f"=B{r}*Настройки!$B$3")
    ws_master.cell(row=r, column=5, value=f"=B{r}*Настройки!$B$4")
    ws_master.cell(row=r, column=6, value=f"=B{r}*Настройки!$B$5")
    ws_master.cell(row=r, column=7, value=(
        f'=SUMIFS(Расходы!$G:$G,Расходы!$B:$B,$B$1,Расходы!$D:$D,$A{r})'
    ))
    ws_master.cell(row=r, column=8, value=f"=G{r}-F{r}")
    ws_master.cell(row=r, column=9, value=f"=E{r}-G{r}")
    ws_master.cell(row=r, column=10, value=f'=IF(B{r}=0,"",I{r}/B{r})')
    ws_master.cell(row=r, column=10).number_format = FMT_MONEY
    ws_master.cell(row=r, column=11, value=f"=B{r}*Настройки!$B$6")
    ws_master.cell(row=r, column=12, value=f"=I{r}-K{r}")
    for col in money_cols_master:
        ws_master.cell(row=r, column=col).number_format = FMT_MONEY

# Тестовые данные: мастер в строке 3
ws_master["A3"] = "Мастер 1"

master_total_row = 101
ws_master.cell(row=master_total_row, column=1, value="Итого")
for col in range(2, 13):
    col_letter = get_column_letter(col)
    cell = ws_master.cell(
        row=master_total_row, column=col,
        value=f"=SUM({col_letter}3:{col_letter}100)",
    )
    if col in money_cols_master or col == 2:
        cell.number_format = FMT_MONEY if col != 2 else FMT_VOLUME
style_total_row(ws_master, master_total_row, 1, 12)

ws_master.column_dimensions["A"].width = 25

freeze_header(ws_master)

# ---------------------------------------------------------------------------
# Лист "Свод_Период"
# ---------------------------------------------------------------------------

period_rows = [
    ("Период", "=Настройки!$B$9", None),
    ("Показатель", "Значение", None),
    ("Объем", "=SUMIF(Распиловка!$B:$B,$B$1,Распиловка!$F:$F)", FMT_VOLUME),
    ("Начислено всего", "=B3*Настройки!$B$2", FMT_MONEY),
    ("Бригаде", "=B3*Настройки!$B$3", FMT_MONEY),
    ("Мастеру валовая", "=B3*Настройки!$B$4", FMT_MONEY),
    ("План расходов", "=B3*Настройки!$B$5", FMT_MONEY),
    ("Факт расходов", "=SUMIF(Расходы!$B:$B,$B$1,Расходы!$G:$G)", FMT_MONEY),
    ("Отклонение расходов", "=B8-B7", FMT_MONEY),
    ("Чистая зарплата мастера", "=B6-B8", FMT_MONEY),
    ("Руб/м3 факт расходов", '=IF(B3=0,"",B8/B3)', FMT_MONEY),
    ("Руб/м3 чистая мастера", '=IF(B3=0,"",B10/B3)', FMT_MONEY),
    ("Проверка: сумма по мастерам", "=SUM(Свод_Мастер!I3:I100)", FMT_MONEY),
    ("Расхождение", "=B10-B13", FMT_MONEY),
]

for i, (label, formula, fmt) in enumerate(period_rows, start=1):
    ws_period.cell(row=i, column=1, value=label)
    cell = ws_period.cell(row=i, column=2, value=formula)
    if fmt:
        cell.number_format = fmt
    for col in (1, 2):
        ws_period.cell(row=i, column=col).border = BORDER

ws_period["A1"].font = HEADER_FONT
ws_period["A2"].font = HEADER_FONT
ws_period["B2"].font = HEADER_FONT
ws_period["A1"].fill = HEADER_FILL
ws_period["A2"].fill = HEADER_FILL
ws_period["B2"].fill = HEADER_FILL

ws_period.column_dimensions["A"].width = 25
ws_period.column_dimensions["B"].width = 18

freeze_header(ws_period)

# ---------------------------------------------------------------------------
# Автоширина столбцов (кроме листов с 10000 строк - там фикс. ширина)
# ---------------------------------------------------------------------------

autofit_columns(ws_settings)
autofit_columns(ws_ref)
autofit_columns(ws_planfact)
autofit_columns(ws_master)
autofit_columns(ws_period)

# Для Распиловка и Расходы задаём ширину по заголовкам (без обхода 10000 строк)
for ws, headers in ((ws_raspil, raspil_header), (ws_rashody, rashody_header)):
    for col, title in enumerate(headers, start=1):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = max(12, min(30, len(title) + 4))

# Ширина столбца A на листах сводов = 25 (переустановим точно, autofit мог изменить)
ws_planfact.column_dimensions["A"].width = 25
ws_master.column_dimensions["A"].width = 25
ws_period.column_dimensions["A"].width = 25

# ---------------------------------------------------------------------------
# Сохранение
# ---------------------------------------------------------------------------

wb.save(OUTPUT_FILE)
print(f"Файл '{OUTPUT_FILE}' создан.")

# ---------------------------------------------------------------------------
# Самопроверка: открываем файл и убеждаемся, что формулы записаны
# ---------------------------------------------------------------------------

check_wb = load_workbook(OUTPUT_FILE)

expected_sheets = [
    "Инструкция", "Настройки", "Справочник_Статей", "Распиловка",
    "Расходы", "ПланФакт_Расходы", "Свод_Мастер", "Свод_Период",
]
assert check_wb.sheetnames == expected_sheets, (
    f"Порядок листов не совпадает: {check_wb.sheetnames}"
)

checks = [
    ("Настройки", "B4", "=B2-B3"),
    ("Настройки", "B6", "=B4-B5"),
    ("Справочник_Статей", "C11", "=SUM(C2:C10)"),
    ("Распиловка", "B2", '=IF(A2="","",TEXT(A2,"MM.YYYY"))'),
    ("Распиловка", "H2", '=IF(F2="","",F2*G2)'),
    ("Расходы", "B2", '=IF(A2="","",TEXT(A2,"MM.YYYY"))'),
    ("ПланФакт_Расходы", "B2", "=SUMIF(Распиловка!$B:$B,$B$1,Распиловка!$F:$F)"),
    ("Свод_Мастер", "B3", '=SUMIFS(Распиловка!$F:$F,Распиловка!$B:$B,$B$1,Распиловка!$D:$D,$A3)'),
    ("Свод_Период", "B3", "=SUMIF(Распиловка!$B:$B,$B$1,Распиловка!$F:$F)"),
]

for sheet_name, cell_ref, expected_formula in checks:
    ws_check = check_wb[sheet_name]
    actual = ws_check[cell_ref].value
    assert actual == expected_formula, (
        f"{sheet_name}!{cell_ref}: ожидалось {expected_formula!r}, получено {actual!r}"
    )

# Проверка тестовых данных
assert check_wb["Распиловка"]["F2"].value == 61.206
assert check_wb["Расходы"]["G2"].value == 10600
assert check_wb["Расходы"]["G3"].value == 7300
assert check_wb["Расходы"]["G4"].value == 3050
assert check_wb["Расходы"]["G5"].value == 4000

# Проверка Data Validation на листе "Расходы"
dvs = check_wb["Расходы"].data_validations.dataValidation
assert any(
    dv.formula1 == "Справочник_Статей!$A$2:$A$10" for dv in dvs
), "Data Validation для столбца Статья не найдена"

# Проверка периода в Настройки!B9
assert check_wb["Настройки"]["B9"].value == "09.2025"

print("Самопроверка пройдена: все формулы, тестовые данные и проверки на месте.")
print(
    "Внимание: openpyxl не вычисляет формулы. Откройте файл в Excel или "
    "LibreOffice Calc — при первом открытии формулы автоматически пересчитаются."
)
