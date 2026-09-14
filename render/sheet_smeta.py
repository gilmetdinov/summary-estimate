"""
Рендер листа отдельной сметы: шапка, таблица позиций, итоги, подписи.
"""
from __future__ import annotations
from copy import copy

from openpyxl.utils import get_column_letter

from smeta_parser import Smeta
from config import Config
from render.donor import Donor
from render.style import appendix, clean_quotes, merge_block, put


def render_smeta(wb, sm: Smeta, d: Donor, title: str, cfg: Config) -> dict:
    """Рисует лист одной сметы. Возвращает {title, subtotal_addr, total_addr}."""
    ws = wb.create_sheet(title=title)
    S = d.sm  # донор-лист сметы
    # ширины колонок A..F (F — служебная, скрытая)
    for col, w in {"A": 7.0, "B": 32.1, "C": 47.1, "D": 40.3, "E": 12.1}.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["F"].hidden = True
    ws.sheet_view.showGridLines = False

    # служебная колонка F: параметры для CONCATENATE-формул.
    # Пишем project_code как строку: int() режет ведущие нули ("04" -> 4),
    # а они должны сохраняться в номере приложения.
    appendix_n = appendix(sm.suffix).rsplit(".", 1)[-1]
    ws.cell(row=1, column=6).value = sm.project_code
    ws.cell(row=6, column=6).value = sm.project_code

    # --- шапка ---
    put(ws, 1, 5, f'=CONCATENATE("Приложение №3.",F1,".","{appendix_n}")', S["E1"], height=12.75)
    put(ws, 2, 5, "к договору", S["E2"], height=12.75)
    put(ws, 3, 5, "от ____.____._______г.", S["E2"], height=12.75)
    put(ws, 4, 5, "№_________________", S["E2"], height=12.75)
    for r in (2, 3, 4):
        c = ws.cell(row=r, column=5)
        a = copy(c.alignment)
        a.wrap_text = False
        c.alignment = a
    put(ws, 6, 1, f'=CONCATENATE("СМЕТА № ",F6,".","{sm.suffix}")', S["A6"], c2=5, height=15.0)
    # скрытая строка между шапкой и таблицей (как в target)
    ws.row_dimensions[17].hidden = True
    put(ws, 7, 1, sm.work_line, S["A7"], c2=5, height=15.0)
    put(ws, 8, 1, clean_quotes(sm.object_full), S["A8"], c2=5, height=25.5)
    put(ws, 9, 1, "(наименование стройки)", S["A9"], c2=5, height=15.0)
    put(ws, 11, 1, "Заказчик", S["A11"], height=15.0)
    put(ws, 12, 2, clean_quotes(sm.customer), S["B12"], c2=5, height=15.0)
    put(ws, 13, 1, "(наименование организации)", S["A13"], c2=5, height=15.0)
    put(ws, 14, 1, "Проектная организация", S["A14"], height=15.0)
    put(ws, 15, 2, clean_quotes(sm.org), S["B15"], c2=5, height=15.0)
    put(ws, 16, 1, "(наименование организации)", S["A16"], c2=5, height=15.0)
    put(ws, 18, 1, sm.price_level, S["A18"], height=15.0)

    # --- шапка таблицы ---
    hdr = ["№ пп",
           "Наименование\nобъекта проектирования или вида проектных работ",
           "Наименование, номера глав, таблиц, параграфов и пунктов МНЗ на проектные работы",
           "Расчет стоимости", "Сметная стоимость, тыс. руб."]
    for c, txt in enumerate(hdr, start=1):
        put(ws, 20, c, txt, S[f"{get_column_letter(c)}20"], height=36.0)
    for c in range(1, 6):
        put(ws, 21, c, c, S["A21"], height=15.0)

    R = 22
    n_items = 0
    for sec in sm.sections:
        put(ws, R, 1, sec.title, S["A22"], c2=5, height=15.0)
        R += 1
        for it in sec.items:
            n_items += 1
            put(ws, R, 1, it.n, S["A23"])
            put(ws, R, 2, it.name, S["B23"])
            put(ws, R, 3, it.osn, S["C23"])
            put(ws, R, 4, it.calc, S["D23"])
            put(ws, R, 5, it.total, S["E23"])
            R += 1
            for sl in it.sublines:
                put(ws, R, 1, None, S["A28"])
                put(ws, R, 2, None, S["B28"])
                put(ws, R, 3, sl.text, S["C28"])
                put(ws, R, 4, None, S["D28"])
                val = sl.value if sl.value is not None else None
                put(ws, R, 5, val, S["E28"])
                R += 1

    # --- итоги (без НДС — НДС только в св.смете) ---
    # колонка D (4) пишется явно: без неё у ячейки нет границ -> «дыра» в рамке таблицы
    sec_kind = sm.sections[0].title.replace("Раздел 1.", "").strip() if sm.sections else ""
    put(ws, R, 1, None, S["A49"]); put(ws, R, 2, f"Итого по разделу 1 {sec_kind}", S["B49"], c2=3, height=15.0); put(ws, R, 4, None, S["D49"]); put(ws, R, 5, sm.subtotal_no_vat, S["E49"], height=15.0)
    R += 1
    put(ws, R, 1, None, S["A50"]); put(ws, R, 2, "Итоги по смете:", S["B50"], c2=3, height=15.0); put(ws, R, 4, None, S["D50"]); put(ws, R, 5, None, S["E51"], height=15.0)
    R += 1
    subtotal_addr = f"E{R}"
    put(ws, R, 1, None, S["A51"]); put(ws, R, 2, f"     Итого Поз. 1-{n_items}", S["B51"], c2=3, height=15.0); put(ws, R, 4, None, S["D51"]); put(ws, R, 5, sm.subtotal_no_vat, S["E51"], height=15.0)
    R += 2

    # --- сумма прописью (без НДС, ссылается на субтотал) ---
    put(ws, R, 1, d.propis_for_smeta(subtotal_addr), S["A55"], c2=5, height=12.75)
    R += 1
    put(ws, R, 1, "(сумма прописью)", S["A56"], c2=5, height=12.75)
    R += 2

    # --- подписи ---
    put(ws, R, 1, "Проектная организация:", S["A57"], height=12.75); R += 1
    dir_block = f"{cfg.org_dir_title1} {cfg.org_dir_title2} {cfg.org_short}"
    merge_block(ws, f"A{R}:B{R+2}", dir_block, S["A58"])
    put(ws, R, 4, cfg.org_dir_name, S["D59"]); R += 2
    put(ws, R, 3, "(подпись)", S["C60"]); put(ws, R, 4, "(инициалы, фамилия)", S["C60"]); R += 2
    put(ws, R, 1, "Главный инженер проекта", S["A62"], c2=2); put(ws, R, 4, cfg.gip_name, S["D59"]); R += 1
    put(ws, R, 3, "(подпись)", S["C60"]); put(ws, R, 4, "(инициалы, фамилия)", S["C60"]); R += 1
    put(ws, R, 1, '" _____ " ________________ 20__ г.', S["A64"], c2=2, height=12.75); R += 1
    put(ws, R, 1, "М.П.", S["A65"], height=12.75); R += 2
    put(ws, R, 1, "Заказчик:", S["A67"], height=12.75); R += 1
    cust_block = f"{cfg.cust_dir_title} {cfg.cust_dir_title2}"
    merge_block(ws, f"A{R}:B{R+1}", cust_block, S["A68"])
    put(ws, R, 4, cfg.cust_dir_name, S["D59"]); R += 1
    put(ws, R, 3, "(подпись)", S["C60"]); put(ws, R, 4, "(инициалы, фамилия)", S["C60"]); R += 2
    put(ws, R, 3, "(подпись)", S["C60"]); put(ws, R, 4, "(инициалы, фамилия)", S["C60"]); R += 1
    put(ws, R, 1, '" _____ " ________________ 20__ г.', S["A64"], c2=2, height=12.75); R += 1
    put(ws, R, 1, "М.П.", S["A65"], height=12.75)

    # печать
    ws.print_area = f"A1:E{R}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.scale = 68
    return {"title": title, "subtotal_addr": subtotal_addr}
