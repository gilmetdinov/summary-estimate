"""
Генератор демо-выгрузки examples/demo_raw.xlsx — ВЫМЫШЛЕННЫЕ данные для примера.

Объект, организации, ФИО и суммы придуманы (круглые числа); совпадения с реальными
сметами случайны. Структура повторяет «сырую» выгрузку сметной программы, которую
понимает smeta_parser.py: шапка -> таблица -> разделы/позиции/подстроки -> итоги -> подписи.

Запуск:  python3 examples/make_demo.py      # перезаписывает examples/demo_raw.xlsx
"""
from __future__ import annotations
import datetime
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DST = os.path.join(HERE, "demo_raw.xlsx")

PROJECT = "99"
OBJECT = "Демо-объект: тепловая магистраль № 1 (вымышленные данные для примера)"
CUSTOMER = "ООО «Демо-Заказчик»"
ORG = "ООО «Демо-Проект»"
PRICE = "Составлена в уровне цен на 1 квартал 2026 г."

# (суффикс, хвост наименования стройки, строка «на ...», разделы)
# раздел: (заголовок, [(наименование, обоснование, расчёт, сумма, [(подстрока, значение)])])
SMETAS = [
    ("01.01", "Проектные работы (проектная документация)", "на проектные работы", [
        ("Раздел 1. Проектные работы (проектная документация)", [
            ("Тепловая сеть, подземная прокладка, 1 км", "Демо-МНЗ, табл. 1, п. 1",
             "1000,0 × 1,0 × 0,4", 400.0,
             [("К = 1,0 (демо-коэффициент)", None),
              ("ПЗ", 100.0), ("ТКР", 200.0), ("ПОС", 50.0), ("СМ", 50.0)]),
            ("Тепловая камера, 2 шт.", "Демо-МНЗ, табл. 2, п. 3",
             "50,0 × 2", 100.0,
             [("ТКР", 80.0), ("СМ", 20.0)]),
        ]),
    ]),
    ("02.01", "Проектные работы (рабочая документация)", "на проектные работы", [
        ("Раздел 1. Проектные работы (рабочая документация)", [
            ("Тепловая сеть, подземная прокладка, 1 км", "Демо-МНЗ, табл. 1, п. 1",
             "1000,0 × 1,0 × 0,6", 600.0,
             [("ТКР", 500.0), ("СМ", 100.0)]),
        ]),
    ]),
    ("03", "Инженерно-геодезические изыскания", "на изыскательские работы", [
        ("Раздел 1. Полевые работы", [
            ("Топографическая съёмка М 1:500, 10 га", "Демо-НЗ_ИГДИ, табл. 1",
             "20,0 × 10", 200.0, []),
        ]),
        ("Раздел 2. Камеральные работы", [
            ("Обработка материалов съёмки, 10 га", "Демо-НЗ_ИГДИ, табл. 2",
             "10,0 × 10", 100.0, []),
        ]),
    ]),
]


def _sheet(wb, suffix, tail, work_line, sections):
    number = f"{PROJECT}.{suffix}"
    ws = wb.create_sheet(f"смета {number}")
    rows = [
        (f"СМЕТА № {number}",),
        (f"{OBJECT}, {tail}",),
        (work_line,),
        (PRICE,),
        ("Заказчик",),
        (None, CUSTOMER),
        ("Проектная организация",),
        (None, ORG),
        ("№ пп", "Наименование", "Обоснование", "Расчет стоимости", "Сметная стоимость, тыс. руб."),
        (1, 2, 3, 4, 5),
    ]
    n = 0
    total = 0.0
    for title, items in sections:
        rows.append((title,))
        for name, osn, calc, value, sublines in items:
            n += 1
            total += value
            rows.append((str(n), name, osn, calc, value))
            for text, v in sublines:
                rows.append((None, None, text, None, v))
    vat = round(total * 0.22, 2)
    rows += [
        (None, "Итого по разделу 1", None, None, total),
        (None, f"Итого Поз. 1-{n}", None, None, total),
        (None, "НДС 22%", None, None, vat),
        (None, "Всего по смете", None, None, round(total + vat, 2)),
        (None, "Руководитель"),
    ]
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            if v is not None:
                ws.cell(row=r, column=c).value = v


def build(dst: str = DST) -> str:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for spec in SMETAS:
        _sheet(wb, *spec)
    wb.properties.creator = "summary-estimate demo"
    wb.properties.lastModifiedBy = None
    wb.properties.created = datetime.datetime(2026, 1, 1)
    wb.save(dst)
    print(f"OK -> {dst}")
    return dst


if __name__ == "__main__":
    build()
