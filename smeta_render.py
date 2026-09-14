"""
Конвертер выгрузки смет ПИР -> фирменный печатный формат (листы смет + сводная).

Единый entrypoint: этот модуль — публичный шлюз. Вся логика рендеринга вынесена
в пакет `render/` (см. render/__init__.py) и разбита по модулям:
- render/convert.py — оркестрация (convert / convert_batch)
- render/sheet_smeta.py / sheet_svodnaya.py / sheet_izyskaniya.py — рендер листов
- render/donor.py — мастер-шаблон оформления
- render/style.py — низкоуровневые хелперы openpyxl

`app.py` (GUI) и сборка PyInstaller импортируют convert/convert_batch/suggest_name
через этот модуль — стабильный путь входа.
"""
from __future__ import annotations

# re-export публичного API для обратной совместимости (app.py, smoke_test.py)
from render import (  # noqa: F401
    convert,
    convert_batch,
    suggest_name,
    default_report_name,
    Donor,
    SMETA_DONOR,
    SVOD_DONOR,
    DEFAULT_DONOR,
    render_smeta,
    render_svodnaya,
    render_izyskaniya_placeholder,
)
from config import load_config  # noqa: F401 — используется CLI-режимом


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(
        description="Конвертер выгрузки смет ПИР -> фирменный печатный формат.")
    ap.add_argument("inputs", nargs="*", help="выгрузка(и) ГС (xlsx)")
    ap.add_argument("-o", "--output", help="файл результата (1 вход) или папка (батч)")
    ap.add_argument("--config", default="config.json", help="config.json с реквизитами")
    ap.add_argument("--donor", default=DEFAULT_DONOR, help="файл-донор оформления (мастер-шаблон)")
    args = ap.parse_args()

    cfg = load_config(args.config)
    inputs = args.inputs
    if not inputs:
        ap.error("укажите хотя бы одну выгрузку ГС (xlsx)")

    # legacy: `smeta_render.py raw.xlsx out.xlsx` (2 позиционных, без -o)
    if args.output is None and len(inputs) == 2 and inputs[1].lower().endswith(".xlsx"):
        convert(inputs[0], inputs[1], args.donor, cfg)
        print(f"OK -> {inputs[1]}")
    elif len(inputs) == 1:
        out = args.output or suggest_name(inputs[0])
        convert(inputs[0], out, args.donor, cfg)
        print(f"OK -> {out}")
    else:
        out_dir = args.output or "converted"
        res = convert_batch(inputs, out_dir, args.donor, cfg)
        ok = sum(1 for r in res if r["error"] is None)
        print(f"\nБатч готов: {ok}/{len(res)} -> {out_dir}/")
        for r in res:
            print(f"  {r['input']} -> {r['output'] or 'ОШИБКА: ' + r['error']}")
