"""
Сборка мастер-шаблона template.xlsx из эталона example.xlsx.

Мастер-шаблон — выделенный файл-донор ОФОРМЛЕНИЯ: только два листа-донора (лист
сметы и лист сводной) под стабильными именами. example.xlsx остаётся эталоном/фикстурой
и от прода отвязан. Рендерер (smeta_render.py) и сборка exe ссылаются на template.xlsx.

ВАЖНО (конфиденциальность): example.xlsx содержит РЕАЛЬНУЮ смету заказчика. Поэтому
после копирования листов-доноров все значения ячеек (данные) ОЧИЩАЮТСЯ — сохраняются
только оформление (стили/рамки/форматы/высоты/объединения) и формулы «сумма прописью».
template.xlsx в репозитории, таким образом, безопасен для публикации.

Запуск:  python3 make_template.py      # перегенерировать чистый шаблон из example.xlsx
"""
from __future__ import annotations
import openpyxl

SRC = "sources/example.xlsx"   # эталон-источник (локальный, под .gitignore)
DST = "template.xlsx"          # мастер-шаблон (коммитится, вкладывается в exe)

# исходные листы-доноры в example.xlsx
SRC_SMETA = "смета 04.01.01_ПД_без БИМ - Фор"
SRC_SVOD = "св.смета_04"

# стабильные имена в мастер-шаблоне (ДОЛЖНЫ совпадать с константами в smeta_render.py)
DONOR_SMETA = "ШАБЛОН_смета"
DONOR_SVOD = "ШАБЛОН_сводная"


def _scrub(wb) -> int:
    """Очищает значения ячеек (реальные данные сметы), сохраняя стили и формулы.

    Формулы (value начинается с '=') оставляем — это «сумма прописью» (A55/C30),
    они нужны рендереру (Donor.propis_smeta / propis_svod).
    """
    cleared = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, str) and v.startswith('='):
                    continue
                cell.value = None
                cleared += 1
    return cleared


def build(src: str = SRC, dst: str = DST) -> str:
    wb = openpyxl.load_workbook(src)
    keep = {SRC_SMETA, SRC_SVOD}
    missing = keep - set(wb.sheetnames)
    if missing:
        raise ValueError(f"В {src!r} нет листов-доноров: {missing}")
    for name in list(wb.sheetnames):
        if name not in keep:
            del wb[name]
    wb[SRC_SMETA].title = DONOR_SMETA
    wb[SRC_SVOD].title = DONOR_SVOD
    cleared = _scrub(wb)
    wb.save(dst)
    print(f"OK -> {dst}  (листы: {wb.sheetnames}; очищено данных: {cleared} ячеек)")
    return dst


if __name__ == "__main__":
    build()

