"""
Пакет рендеринга смет ПИР.

Публичный API (используется `smeta_render.py` и `app.py`):
- convert / convert_batch — оркестрация (convert.py)
- suggest_name / default_report_name — имена файлов (convert.py)
- render_smeta / render_svodnaya / render_izyskaniya_placeholder — листы (sheet_*.py)
- Donor + имена листов-доноров (donor.py)
"""
from render.donor import Donor, SMETA_DONOR, SVOD_DONOR, DEFAULT_DONOR
from render.convert import (
    convert,
    convert_batch,
    suggest_name,
    default_report_name,
)
from render.sheet_smeta import render_smeta
from render.sheet_svodnaya import render_svodnaya
from render.sheet_izyskaniya import render_izyskaniya_placeholder

__all__ = [
    "Donor",
    "SMETA_DONOR",
    "SVOD_DONOR",
    "DEFAULT_DONOR",
    "convert",
    "convert_batch",
    "suggest_name",
    "default_report_name",
    "render_smeta",
    "render_svodnaya",
    "render_izyskaniya_placeholder",
]
