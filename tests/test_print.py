"""Печать результата: A4, «вписать по ширине», область печати сводной = колонки формы.

На демо-выгрузке examples/demo_raw.xlsx (вымышленные данные). Проверка пересчёта
формул и числа страниц — через LibreOffice headless, если он установлен (иначе skip).
"""
from __future__ import annotations

import os
import re
import subprocess

import openpyxl
import pytest
from openpyxl.utils import range_boundaries

from config import Config, Placeholder
from render import convert
from render.donor import DEFAULT_DONOR
from smoke_test import find_soffice

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEMO = os.path.join(ROOT, "examples", "demo_raw.xlsx")
ERRORS = ("#DIV/0!", "#REF!", "#NAME?", "#VALUE!", "#N/A", "#NUM!", "#NULL!")

pytestmark = pytest.mark.skipif(
    not (os.path.exists(DEFAULT_DONOR) and os.path.exists(DEMO)),
    reason="нет template.xlsx или examples/demo_raw.xlsx")


def _cfg_full() -> Config:
    """Конфиг со строками изысканий и госэкспертизы — покрывает лист-заготовку изысканий."""
    return Config(
        izyskaniya=Placeholder(name="Демо-изыскания", osn="Смета №99.03", value=100.0, column="izysk"),
        gosexpertiza=Placeholder(name="Демо-госэкспертиза", osn="Расчёт", value=50.0, column="proj"),
    )


@pytest.fixture(params=["default", "full"])
def result(request, tmp_path):
    cfg = Config() if request.param == "default" else _cfg_full()
    out = tmp_path / f"demo_{request.param}.xlsx"
    convert(DEMO, str(out), DEFAULT_DONOR, cfg)
    return str(out)


def test_every_sheet_is_a4_fit_to_width(result):
    wb = openpyxl.load_workbook(result)
    for ws in wb.worksheets:
        ps = ws.page_setup
        assert int(ps.paperSize) == 9, ws.title  # 9 = A4
        assert ps.orientation == "portrait", ws.title
        assert int(ps.fitToWidth) == 1, ws.title
        assert int(ps.fitToHeight) == 0, ws.title
        assert ws.sheet_properties.pageSetUpPr.fitToPage, ws.title
        assert ws.page_margins.right < 0.5, ws.title  # поля из донора, не дефолтные 0.75"


def test_summary_print_area_is_form_columns(result):
    ws = openpyxl.load_workbook(result)["св.смета"]
    area = ws.print_area.split("!")[-1].replace("$", "")
    min_col, _, max_col, _ = range_boundaries(area)
    assert (min_col, max_col) == (1, 7)  # A:G — служебные H/I/O/P вне печати


def test_unit_cost_division_is_guarded(result):
    ws = openpyxl.load_workbook(result)["св.смета"]
    divs = [c.value for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and "/O" in c.value]
    assert divs
    assert all(v.startswith("=IFERROR(") for v in divs)


def test_survey_sheet_header_matches_kind(result):
    wb = openpyxl.load_workbook(result)
    for ws in wb.worksheets:
        if ws.title == "св.смета":
            continue
        header = " ".join(str(ws.cell(row=20, column=c).value) for c in (2, 3))
        survey = ws.title.endswith(".03") or ws.title.startswith("изыскания")
        assert ("изыскательск" in header) == survey, ws.title
        assert ("проектных работ" in header) != survey, ws.title


def _soffice_convert(src: str, fmt: str, outdir: str, soffice: str) -> str:
    # отдельный профиль: не конфликтует с уже запущенным LibreOffice
    profile = "file://" + os.path.join(outdir, "lo-profile")
    subprocess.run([soffice, f"-env:UserInstallation={profile}", "--headless",
                    "--convert-to", fmt, "--outdir", outdir, src],
                   check=True, capture_output=True, timeout=180)
    out = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + "." + fmt)
    assert os.path.exists(out)
    return out


@pytest.mark.skipif(find_soffice() is None, reason="LibreOffice не установлен")
def test_no_error_values_after_recalc(result, tmp_path):
    recalced = _soffice_convert(result, "xlsx", str(tmp_path / "recalc"), find_soffice())
    wb = openpyxl.load_workbook(recalced, data_only=True)
    bad = [f"{ws.title}!{c.coordinate}={c.value}" for ws in wb.worksheets
           for row in ws.iter_rows() for c in row
           if isinstance(c.value, str) and c.value.strip() in ERRORS]
    assert not bad


@pytest.mark.skipif(find_soffice() is None, reason="LibreOffice не установлен")
def test_pdf_one_page_per_sheet(result, tmp_path):
    """Каждый лист демо-комплекта — ровно одна страница A4 (колонки не уезжают)."""
    pdf = _soffice_convert(result, "pdf", str(tmp_path / "pdf"), find_soffice())
    data = open(pdf, "rb").read()
    pages = len(re.findall(rb"/Type\s*/Page(?!s)", data))
    n_sheets = len(openpyxl.load_workbook(result).sheetnames)
    assert pages == n_sheets
    sizes = re.findall(rb"/MediaBox\s*\[\s*0 0 ([\d.]+) ([\d.]+)", data)
    assert sizes and all(abs(float(w) - 595.3) < 1 and abs(float(h) - 841.9) < 1 for w, h in sizes)
