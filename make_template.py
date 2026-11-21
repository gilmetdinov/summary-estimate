"""
Сборка мастер-шаблона template.xlsx из эталона example.xlsx.

Мастер-шаблон — выделенный файл-донор ОФОРМЛЕНИЯ: только два листа-донора (лист
сметы и лист сводной) под стабильными именами. example.xlsx остаётся эталоном/фикстурой
и от прода отвязан. Рендерер (smeta_render.py) и сборка exe ссылаются на template.xlsx.

ВАЖНО (конфиденциальность): example.xlsx содержит РЕАЛЬНУЮ смету заказчика. Поэтому
после копирования листов-доноров очищаются ВСЕ значения и формулы, кроме формул
«сумма прописью», которые читает рендерер (render/donor.py), а также свойства документа
(автор, кем изменён и т.п.), именованные диапазоны, внешние ссылки, примечания,
гиперссылки и колонтитулы. Сохраняется только оформление (стили/рамки/форматы/высоты/
объединения/параметры печати). template.xlsx в репозитории, таким образом, безопасен
для публикации.

Запуск:
    python3 make_template.py                  # из sources/example.xlsx (локальный эталон)
    python3 make_template.py template.xlsx    # пере-очистить уже готовый шаблон на месте
    python3 make_template.py эталон.xlsx --smeta-sheet "<лист сметы>" --svod-sheet "<лист сводной>"
"""
from __future__ import annotations
import argparse
import datetime
import os

import openpyxl
from openpyxl.packaging.core import DocumentProperties
from openpyxl.worksheet.header_footer import HeaderFooter

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "sources", "example.xlsx")   # эталон-источник (локальный, под .gitignore)
DST = os.path.join(HERE, "template.xlsx")             # мастер-шаблон (коммитится, вкладывается в exe)

# листы-доноры в эталоне: по умолчанию — первый лист, имя которого начинается с префикса
# (точные имена листов эталона в код не зашиваем; при неоднозначности — --smeta-sheet/--svod-sheet)
SRC_SMETA_PREFIX = "смета"
SRC_SVOD_PREFIX = "св.смета"

# стабильные имена в мастер-шаблоне (ДОЛЖНЫ совпадать с константами в render/donor.py)
DONOR_SMETA = "ШАБЛОН_смета"
DONOR_SVOD = "ШАБЛОН_сводная"

# единственные формулы, которые нужны рендереру: «сумма прописью» (Donor.propis_smeta /
# propis_svod). Всё остальное в эталоне — ссылки на листы исходной книги и данные.
KEEP_FORMULAS = {DONOR_SMETA: {"A55"}, DONOR_SVOD: {"C30"}}

# нейтральные свойства документа (вместо автора/компьютера исходного файла)
NEUTRAL_CREATOR = "summary-estimate"
NEUTRAL_DATE = datetime.datetime(2026, 1, 1)


def _scrub(wb) -> int:
    """Очищает значения и формулы ячеек, кроме KEEP_FORMULAS; стили сохраняются."""
    cleared = 0
    for ws in wb.worksheets:
        keep = KEEP_FORMULAS.get(ws.title, set())
        for row in ws.iter_rows():
            for cell in row:
                if cell.comment is not None:
                    cell.comment = None
                if cell.hyperlink is not None:
                    cell.hyperlink = None
                v = cell.value
                if v is None:
                    continue
                if cell.coordinate in keep and isinstance(v, str) and v.startswith("="):
                    continue
                cell.value = None
                cleared += 1
        # именованные диапазоны листа, автофильтр (_FilterDatabase), колонтитулы
        ws.defined_names.clear()
        ws.auto_filter.ref = None
        ws.HeaderFooter = HeaderFooter()
    return cleared


def _neutralize(wb) -> None:
    """Свойства документа, книжные имена и внешние ссылки -> нейтральные/пустые."""
    wb.properties = DocumentProperties(creator=NEUTRAL_CREATOR, lastModifiedBy=None,
                                       created=NEUTRAL_DATE, modified=NEUTRAL_DATE)
    wb.defined_names.clear()
    wb._external_links = []


def _pick(names: list[str], prefix: str, explicit: str | None) -> str:
    if explicit:
        if explicit not in names:
            raise ValueError(f"нет листа {explicit!r}; есть: {names}")
        return explicit
    for n in names:
        if n.lower().startswith(prefix):
            return n
    raise ValueError(f"нет листа с префиксом {prefix!r}; есть: {names} "
                     f"(укажите --smeta-sheet/--svod-sheet)")


def build(src: str = SRC, dst: str = DST, smeta_sheet: str | None = None,
          svod_sheet: str | None = None) -> str:
    wb = openpyxl.load_workbook(src)
    names = wb.sheetnames
    if {DONOR_SMETA, DONOR_SVOD} <= set(names):
        # уже мастер-шаблон (повторная очистка на месте)
        src_smeta, src_svod = DONOR_SMETA, DONOR_SVOD
    else:
        # «св.смета...» не начинается с «смета», так что префиксы не пересекаются
        src_svod = _pick(names, SRC_SVOD_PREFIX, svod_sheet)
        src_smeta = _pick(names, SRC_SMETA_PREFIX, smeta_sheet)
    for name in list(names):
        if name not in (src_smeta, src_svod):
            del wb[name]
    wb[src_smeta].title = DONOR_SMETA
    wb[src_svod].title = DONOR_SVOD
    cleared = _scrub(wb)
    _neutralize(wb)
    wb.save(dst)
    print(f"OK -> {dst}  (листы: {wb.sheetnames}; очищено: {cleared} ячеек)")
    return dst


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Сборка/очистка мастер-шаблона template.xlsx")
    ap.add_argument("src", nargs="?", default=SRC,
                    help="эталон (по умолчанию sources/example.xlsx) или готовый template.xlsx")
    ap.add_argument("-o", "--output", default=DST, help="куда сохранить (по умолчанию template.xlsx)")
    ap.add_argument("--smeta-sheet", help="точное имя листа сметы в эталоне")
    ap.add_argument("--svod-sheet", help="точное имя листа сводной в эталоне")
    a = ap.parse_args()
    build(a.src, a.output, a.smeta_sheet, a.svod_sheet)
