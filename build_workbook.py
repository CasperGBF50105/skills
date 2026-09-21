# -*- coding: utf-8 -*-
"""
Генератор Excel-файла "Учет_вагонов.xlsx".

Листы:
  1. Параметры   — глобальные константы (курс, цены, транзит и т.д.)
  2. Курсы       — таблица tblКурсы (Дата, Курс USD/RUB)
  3. Вагоны      — таблица tblВагоны (одна строка = один вагон)
  4. Платежи     — таблица tblПлатежи (одна строка = один платёж)
  5. План оплат  — таблица tblПланОплат (график будущих поступлений)
  6. Дашборд     — сводка по всем вагонам

Запуск:
    pip install openpyxl
    python build_workbook.py
"""

from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

OUTPUT = "Учет_вагонов.xlsx"

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

MONEY_FMT = "# ##0.00"
DATE_FMT = "DD.MM.YYYY"
RATE_FMT = "0.00"


def style_header(ws, row=1, ncols=None):
    ncols = ncols or ws.max_column
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN


def autosize(ws, min_w=12, max_w=30):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        longest = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[letter].width = max(min_w, min(max_w, longest + 2))


def add_table(ws, name, ref):
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tbl)


wb = Workbook()

# ============================================================
# 1. Параметры
# ============================================================
ws = wb.active
ws.title = "Параметры"
ws.append(["Параметр", "Значение", "Ед."])
ws.append(["Курс USD/RUB текущий", 91, "RUB/USD"])
ws.append(["Цена поставщика с доставкой", 13700, "RUB/м3"])
ws.append(["Погрузка вагона с дорогой", 720000, "RUB/вагон"])
ws.append(["Транзит по Азербайджану", 2250, "USD/вагон"])
ws.append(["Обратка порожнего вагона", 250, "USD/вагон"])
ws.append(["Валюта цены продажи", "USD", "—"])

style_header(ws)
ws["B2"].number_format = RATE_FMT
ws["B3"].number_format = MONEY_FMT
ws["B4"].number_format = MONEY_FMT
ws["B5"].number_format = MONEY_FMT
ws["B6"].number_format = MONEY_FMT
autosize(ws, min_w=14, max_w=32)
ws.freeze_panes = "A2"

# ============================================================
# 2. Курсы
# ============================================================
ws = wb.create_sheet("Курсы")
ws.append(["Дата", "Курс USD/RUB"])
ws.append([date(2026, 1, 1), 91])
ws.append([date(2026, 1, 15), 92])
ws.append([date(2026, 2, 1), 90])
style_header(ws)
for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=1).number_format = DATE_FMT
    ws.cell(row=r, column=2).number_format = RATE_FMT
autosize(ws, min_w=14, max_w=20)
ws.freeze_panes = "A2"

add_table(ws, "tblКурсы", f"A1:B{ws.max_row}")

# ============================================================
# 3. Вагоны
# ============================================================
ws = wb.create_sheet("Вагоны")

headers_vag = [
    "ID вагона",
    "Номер вагона",
    "Дата отправки",
    "Дата прибытия на ст. Ирана",
    "Дата фиксации цены",
    "Объём, м3",
    "Цена продажи, USD/м3",
    "Курс на дату фиксации",
    "Сумма к оплате, USD",
    "Сумма к оплате, RUB на дату фиксации",
    "Оплачено, USD",
    "Оплачено, RUB",
    "Остаток, USD",
    "Остаток, RUB по текущему курсу",
    "Статус",
    "Расход поставщика, RUB",
    "Погрузка, RUB",
    "Транзит+обратка, USD",
    "Транзит+обратка, RUB",
    "Начисленная прибыль, RUB",
    "Кассовая прибыль, RUB",
    "Курсовая разница, RUB",
    "Прибыль на 1 м3, RUB",
    "Последний платёж",
    "Срок оплаты",
    "Просрочка",
    "Комментарий",
]
ws.append(headers_vag)

# Пример данных
sample_vag = [
    ["В-001", "12345678", date(2026, 1, 5), date(2026, 1, 20), date(2026, 1, 20), 70, 200, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, date(2026, 2, 28), None, ""],
    ["В-002", "23456789", date(2026, 1, 18), date(2026, 2, 2), date(2026, 2, 2), 68, 205, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, date(2026, 3, 15), None, ""],
    ["В-003", "34567890", date(2026, 2, 1), date(2026, 2, 18), date(2026, 2, 18), 72, 210, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, date(2026, 3, 30), None, ""],
]
for row in sample_vag:
    ws.append(row)

style_header(ws)

# Формулы для каждой строки данных
#
# ВАЖНО: OOXML всегда хранит формулы с каноническими английскими именами
# функций (IF/SUM/SUMIFS/...) и запятой как разделителем аргументов —
# независимо от локали. Excel и LibreOffice Calc сами отображают их
# локализованными (ЕСЛИ/СУММ/СУММЕСЛИМН/... с разделителем ";") при
# открытии файла, но в самом файле должны быть английские имена и запятые,
# иначе формулы будут считаться повреждёнными.
for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=8).value = (
        "=XLOOKUP([@[Дата фиксации цены]],tblКурсы[Дата],tblКурсы[Курс],,-1)"
    )
    ws.cell(row=r, column=9).value = (
        "=[@[Объём, м3]]*[@[Цена продажи, USD/м3]]"
    )
    ws.cell(row=r, column=10).value = (
        "=[@[Сумма к оплате, USD]]*[@[Курс на дату фиксации]]"
    )
    ws.cell(row=r, column=11).value = (
        '=SUMIFS(tblПлатежи[Сумма в USD],tblПлатежи[ID вагона],[@[ID вагона]],tblПлатежи[Статус],"факт")'
    )
    ws.cell(row=r, column=12).value = (
        '=SUMIFS(tblПлатежи[Сумма в RUB],tblПлатежи[ID вагона],[@[ID вагона]],tblПлатежи[Статус],"факт")'
    )
    ws.cell(row=r, column=13).value = (
        "=[@[Сумма к оплате, USD]]-[@[Оплачено, USD]]"
    )
    ws.cell(row=r, column=14).value = (
        "=[@[Остаток, USD]]*Параметры!$B$2"
    )
    ws.cell(row=r, column=15).value = (
        '=IF([@[Дата прибытия на ст. Ирана]]="","Ожидает прибытия",'
        'IF([@[Остаток, USD]]<=0,"Оплачен",'
        'IF([@[Оплачено, USD]]>0,"Частично оплачен","Ожидает оплаты")))'
    )
    ws.cell(row=r, column=16).value = (
        "=[@[Объём, м3]]*Параметры!$B$3"
    )
    ws.cell(row=r, column=17).value = "=Параметры!$B$4"
    ws.cell(row=r, column=18).value = "=Параметры!$B$5+Параметры!$B$6"
    ws.cell(row=r, column=19).value = (
        "=[@[Транзит+обратка, USD]]*[@[Курс на дату фиксации]]"
    )
    ws.cell(row=r, column=20).value = (
        "=[@[Сумма к оплате, RUB на дату фиксации]]"
        "-[@[Расход поставщика, RUB]]"
        "-[@[Погрузка, RUB]]"
        "-[@[Транзит+обратка, RUB]]"
    )
    ws.cell(row=r, column=21).value = (
        "=[@[Оплачено, RUB]]"
        "-[@[Расход поставщика, RUB]]"
        "-[@[Погрузка, RUB]]"
        "-[@[Транзит+обратка, RUB]]"
    )
    ws.cell(row=r, column=22).value = (
        "=[@[Оплачено, RUB]]-[@[Оплачено, USD]]*[@[Курс на дату фиксации]]"
    )
    ws.cell(row=r, column=23).value = (
        "=IF([@[Объём, м3]]=0,0,[@[Кассовая прибыль, RUB]]/[@[Объём, м3]])"
    )
    ws.cell(row=r, column=24).value = (
        '=IFERROR(MAXIFS(tblПлатежи[Дата платежа],tblПлатежи[ID вагона],[@[ID вагона]],tblПлатежи[Статус],"факт"),"")'
    )
    ws.cell(row=r, column=26).value = (
        '=IF(AND([@[Срок оплаты]]<>"",[@[Остаток, USD]]>0,TODAY()>[@[Срок оплаты]]),"Просрочено","")'
    )

# Форматы
for r in range(2, ws.max_row + 1):
    for col in (3, 4, 5, 25):
        ws.cell(row=r, column=col).number_format = DATE_FMT
    for col in (6, 7, 8):
        ws.cell(row=r, column=col).number_format = MONEY_FMT
    for col in (9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 22, 23):
        ws.cell(row=r, column=col).number_format = MONEY_FMT
    ws.cell(row=r, column=24).number_format = DATE_FMT

autosize(ws, min_w=12, max_w=26)
ws.freeze_panes = "B2"

add_table(ws, "tblВагоны", f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")

# ============================================================
# 4. Платежи
# ============================================================
ws = wb.create_sheet("Платежи")

headers_pay = [
    "Дата платежа",
    "ID вагона",
    "Сумма платежа",
    "Валюта",
    "Курс на дату платежа",
    "Сумма в USD",
    "Сумма в RUB",
    "Статус",
    "Банк/примечание",
    "№ платежа",
]
ws.append(headers_pay)

sample_pay = [
    [date(2026, 2, 10), "В-001", 15000, "USD", None, None, None, "факт", "Тинькофф", None],
    [date(2026, 2, 25), "В-001", 1000000, "RUB", None, None, None, "факт", "Сбер", None],
    [date(2026, 3, 5), "В-002", 8000, "USD", None, None, None, "факт", "Тинькофф", None],
    [date(2026, 3, 20), "В-002", 600000, "RUB", None, None, None, "факт", "ВТБ", None],
    [date(2026, 4, 10), "В-003", 10000, "USD", None, None, None, "план", "", None],
]
for row in sample_pay:
    ws.append(row)

style_header(ws)

for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=5).value = (
        "=XLOOKUP([@[Дата платежа]],tblКурсы[Дата],tblКурсы[Курс],,-1)"
    )
    ws.cell(row=r, column=6).value = (
        '=IF([@Валюта]="USD",[@[Сумма платежа]],[@[Сумма платежа]]/[@[Курс на дату платежа]])'
    )
    ws.cell(row=r, column=7).value = (
        '=IF([@Валюта]="RUB",[@[Сумма платежа]],[@[Сумма платежа]]*[@[Курс на дату платежа]])'
    )
    ws.cell(row=r, column=10).value = (
        "=COUNTIF($B$2:[@[ID вагона]],[@[ID вагона]])"
    )

for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=1).number_format = DATE_FMT
    ws.cell(row=r, column=3).number_format = MONEY_FMT
    ws.cell(row=r, column=5).number_format = RATE_FMT
    ws.cell(row=r, column=6).number_format = MONEY_FMT
    ws.cell(row=r, column=7).number_format = MONEY_FMT

autosize(ws, min_w=12, max_w=24)
ws.freeze_panes = "A2"

add_table(ws, "tblПлатежи", f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")

# ============================================================
# 5. План оплат
# ============================================================
ws = wb.create_sheet("План оплат")

headers_plan = [
    "ID вагона",
    "№ платежа",
    "Плановая дата",
    "Плановая сумма USD",
    "План RUB",
    "Статус",
    "Факт дата",
    "Факт RUB",
    "Отклонение",
]
ws.append(headers_plan)

sample_plan = [
    ["В-001", 1, date(2026, 2, 10), 15000, None, "план", None, None, None],
    ["В-001", 2, date(2026, 2, 25), 16050, None, "план", None, None, None],
]
for row in sample_plan:
    ws.append(row)

style_header(ws)

for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=9).value = (
        "=IFERROR([@[Факт RUB]]-[@[План RUB]],0)"
    )

for r in range(2, ws.max_row + 1):
    ws.cell(row=r, column=3).number_format = DATE_FMT
    ws.cell(row=r, column=7).number_format = DATE_FMT
    ws.cell(row=r, column=4).number_format = MONEY_FMT
    ws.cell(row=r, column=5).number_format = MONEY_FMT
    ws.cell(row=r, column=8).number_format = MONEY_FMT
    ws.cell(row=r, column=9).number_format = MONEY_FMT

autosize(ws, min_w=12, max_w=22)
ws.freeze_panes = "A2"

add_table(ws, "tblПланОплат", f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")

# ============================================================
# 6. Дашборд
# ============================================================
ws = wb.create_sheet("Дашборд")

ws["A1"] = "Показатель"
ws["B1"] = "Значение"

dash_rows = [
    ("Всего вагонов", "=COUNTA(tblВагоны[ID вагона])"),
    ("Объём, м3", "=SUM(tblВагоны[Объём, м3])"),
    ("Сумма продаж, USD", "=SUM(tblВагоны[Сумма к оплате, USD])"),
    ("Оплачено, USD", "=SUM(tblВагоны[Оплачено, USD])"),
    ("Остаток, USD", "=SUM(tblВагоны[Остаток, USD])"),
    ("Оплачено, RUB", "=SUM(tblВагоны[Оплачено, RUB])"),
    ("Начисленная прибыль, RUB", "=SUM(tblВагоны[Начисленная прибыль, RUB])"),
    ("Кассовая прибыль, RUB", "=SUM(tblВагоны[Кассовая прибыль, RUB])"),
    ("Курсовая разница, RUB", "=SUM(tblВагоны[Курсовая разница, RUB])"),
]

for i, (label, formula) in enumerate(dash_rows, start=2):
    ws.cell(row=i, column=1, value=label)
    ws.cell(row=i, column=2, value=formula)

style_header(ws)
for r in range(2, len(dash_rows) + 2):
    ws.cell(row=r, column=2).number_format = MONEY_FMT

# Заготовки сводных таблиц
ws["A13"] = "Сводка по вагонам"
ws["A14"] = "ID вагона"
ws["B14"] = "Сумма USD"
ws["C14"] = "Оплачено USD"
ws["D14"] = "Остаток USD"
ws["E14"] = "Оплачено RUB"
ws["F14"] = "Кассовая прибыль RUB"
for c in "ABCDEF":
    ws[f"{c}14"].font = HEADER_FONT
    ws[f"{c}14"].fill = HEADER_FILL
    ws[f"{c}14"].alignment = HEADER_ALIGN

ws["A18"] = "Сводка по месяцам"
ws["A19"] = "Месяц"
ws["B19"] = "Оплачено RUB"
ws["C19"] = "Кол-во платежей"
for c in "ABC":
    ws[f"{c}19"].font = HEADER_FONT
    ws[f"{c}19"].fill = HEADER_FILL
    ws[f"{c}19"].alignment = HEADER_ALIGN

ws["A24"] = "Сводка по валютам"
ws["A25"] = "Валюта"
ws["B25"] = "Сумма"
for c in "AB":
    ws[f"{c}25"].font = HEADER_FONT
    ws[f"{c}25"].fill = HEADER_FILL
    ws[f"{c}25"].alignment = HEADER_ALIGN

ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 18
ws.column_dimensions["D"].width = 18
ws.column_dimensions["E"].width = 18
ws.column_dimensions["F"].width = 22
ws.freeze_panes = "A2"

# ============================================================
# Сохранение
# ============================================================
wb.save(OUTPUT)
print(f"Готово: {OUTPUT}")
