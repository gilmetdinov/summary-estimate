"""
Оркестрация конвертации: выгрузка(и) -> готовый xlsx (листы смет + сводная).

Собирает workbook: для каждой распознанной сметы — свой лист (render_smeta), опционально
лист-заготовка изысканий, в начало — сводная со ссылками (render_svodnaya).
"""
from __future__ import annotations
import os

import openpyxl

from smeta_parser import Smeta, parse_workbook
from config import Config
from render.donor import DEFAULT_DONOR, Donor
from render.sheet_smeta import render_smeta
from render.sheet_svodnaya import render_svodnaya
from render.sheet_izyskaniya import render_izyskaniya_placeholder
from render.style import safe_filename, sheet_title, unique_name


def default_report_name(smetas: list[Smeta]) -> str:
    """Имя итогового файла из содержимого выгрузки: 'Сводник_<проект>_<объект>.xlsx'.

    Сметчице сразу понятно, что за отчёт, без ручного переименования. Наименование
    объекта длинное — ограничиваем ~50 символами и режем по границе слова.
    """
    sm = smetas[0]
    proj = (sm.project_code or "").strip()
    obj = safe_filename((sm.object_base or "").strip())
    if len(obj) > 50:
        cut = obj[:50].rsplit(" ", 1)[0].rstrip(" _.,-—«»")
        obj = cut or obj[:50]
    tokens = ["Сводник"]
    if proj:
        tokens.append(proj)
    if obj:
        tokens.append(obj)
    return safe_filename("_".join(tokens))[:120].rstrip(" _.") + ".xlsx"


def suggest_name(in_path: str) -> str:
    """Имя отчёта по выгрузке; при сбое парсинга — запасное от имени входа."""
    try:
        smetas = parse_workbook(in_path)
        if smetas:
            return default_report_name(smetas)
    except Exception:  # noqa: BLE001 — имя не критично, не валим конвертацию
        pass
    stem = os.path.splitext(os.path.basename(in_path))[0]
    return f"{stem}_converted.xlsx"


def convert(in_path: str, out_path: str, donor_path: str = DEFAULT_DONOR,
            cfg: Config | None = None, smetas: list[Smeta] | None = None):
    cfg = cfg or Config()
    if smetas is None:
        smetas = parse_workbook(in_path)
    if not smetas:
        raise ValueError(f"В файле {in_path!r} не распознано ни одной сметы")
    d = Donor(donor_path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used: set = set()
    refs = []
    for sm in smetas:
        title = sheet_title(sm.sheet or f"смета {sm.number}", used)
        refs.append(render_smeta(wb, sm, d, title, cfg))
    if cfg.izyskaniya:
        render_izyskaniya_placeholder(wb, d, cfg, used, smetas[0].project_code)
    render_svodnaya(wb, smetas, refs, d, cfg)
    wb.save(out_path)
    return out_path


def convert_batch(in_paths: list[str], out_dir: str, donor_path: str = DEFAULT_DONOR,
                  cfg: Config | None = None) -> list[dict]:
    """Несколько выгрузок -> папка с результатами за один прогон (один конфиг на все).

    Возвращает список {input, output, error}; error=None при успехе.
    """
    os.makedirs(out_dir, exist_ok=True)
    results = []
    used: set = set()
    for p in in_paths:
        try:
            smetas = parse_workbook(p)
            if not smetas:
                raise ValueError("не распознано ни одной сметы")
            name = unique_name(default_report_name(smetas), used)
            out = os.path.join(out_dir, name)
            convert(p, out, donor_path, cfg, smetas=smetas)
            results.append({"input": p, "output": out, "error": None})
        except Exception as ex:  # noqa: BLE001 — батч не должен падать целиком из-за одного файла
            results.append({"input": p, "output": None, "error": str(ex)})
    return results
