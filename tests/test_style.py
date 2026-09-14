"""Тесты низкоуровневых хелперов рендеринга (render/style.py)."""
from __future__ import annotations

from render.style import (
    appendix,
    clean_quotes,
    safe_filename,
    sheet_title,
    unique_name,
)


def test_appendix_mapping():
    assert appendix("01.01") == "1.1"
    assert appendix("03") == "3"


def test_clean_quotes():
    assert clean_quotes("«название»") == '"название"'
    assert clean_quotes("") == ""


def test_safe_filename_strips_illegal_chars():
    assert safe_filename('a<b>:c?') == "abc"
    assert safe_filename("") == "Сводник"
    assert safe_filename("  x  ") == "x"


def test_sheet_title_unique():
    used = set()
    t1 = sheet_title("смета 12.01.01", used)
    t2 = sheet_title("смета 12.01.01", used)
    assert t1 != t2
    assert len(t1) <= 31
    assert len(t2) <= 31


def test_unique_name_suffixes():
    used = set()
    n1 = unique_name("report.xlsx", used)
    n2 = unique_name("report.xlsx", used)
    assert n1 == "report.xlsx"
    assert n2 == "report_2.xlsx"
