"""Юнит-тесты парсера выгрузки смет (smeta_parser)."""
from __future__ import annotations

import pytest

from smeta_parser import (
    SUFFIX_KIND,
    _num,
    _txt,
    _split_object,
    parse_sheet,
    parse_workbook,
)
from conftest import make_sheet


# --- хелперы ---

def test_num_parses_russian_decimal():
    assert _num("4,08807") == pytest.approx(4.08807)
    assert _num("1\xa0101,2819") == pytest.approx(1101.2819)
    assert _num(" 12,5 ") == pytest.approx(12.5)


def test_num_edge_cases():
    assert _num(None) is None
    assert _num("") is None
    assert _num("abc") is None
    assert _num(42) == 42.0


def test_txt_strips():
    assert _txt("  x  ") == "x"
    assert _txt(None) is None


def test_split_object_strips_kind_tail():
    base, full = _split_object("Стройка, Проектные работы (проектная документация)")
    assert base == "Стройка"
    assert full.startswith("Стройка,")


def test_split_object_no_tail():
    base, full = _split_object("Просто объект")
    assert base == "Просто объект"
    assert full == "Просто объект"


def test_suffix_kind_mapping():
    assert SUFFIX_KIND["01.01"].startswith("Проектные работы")
    assert SUFFIX_KIND["03"] == "Инженерно-геодезические изыскания"


# --- parse_sheet ---

def test_parse_sheet_header():
    ws = make_sheet([
        ("СМЕТА № 12.01.01", None, None, None, None),
        ("Очень длинное наименование объекта строительства которое точно длиннее шестидесяти символов", None, None, None, None),
        ("на проектные работы", None, None, None, None),
        ("Составлена в уровне цен на 1 квартал 2025 г.", None, None, None, None),
        ("Заказчик", None, None, None, None),
        (None, "ООО «Заказчик»", None, None, None),
        ("Проектная организация", None, None, None, None),
        (None, "ООО «Проектная организация»", None, None, None),
        ("№ пп", None, None, None, None),
        ("1", None, None, None, None),
    ])
    sm = parse_sheet(ws)
    assert sm.number == "12.01.01"
    assert sm.project_code == "12"
    assert sm.suffix == "01.01"
    assert sm.kind.startswith("Проектные работы")
    assert sm.customer == "ООО «Заказчик»"
    assert sm.org == "ООО «Проектная организация»"
    assert sm.work_line == "на проектные работы"
    assert sm.price_level.startswith("Составлена в уровне цен")


def test_parse_sheet_body_and_totals(raw_xlsx):
    smetas = parse_workbook(raw_xlsx)
    assert len(smetas) == 1
    sm = smetas[0]

    assert len(sm.sections) == 1
    sec = sm.sections[0]
    assert sec.title == "Раздел 1. Проектные работы"
    assert len(sec.items) == 2

    it = sec.items[0]
    assert it.n == "1"
    assert it.name == "Позиция один"
    assert it.total == pytest.approx(100.5)
    assert len(it.sublines) == 1
    assert it.sublines[0].text == "коэффициент 0,9"
    assert it.sublines[0].value == pytest.approx(12.5)

    assert sm.subtotal_no_vat == pytest.approx(163.0)
    assert sm.vat == pytest.approx(35.86)
    assert sm.grand_total == pytest.approx(198.86)


def test_parse_sheet_missing_header_raises():
    ws = make_sheet([
        ("СМЕТА № 12.01.01", None, None, None, None),
    ])
    with pytest.raises(ValueError):
        parse_sheet(ws)


def test_parse_workbook_skips_bad_sheet(tmp_path):
    """Лист без шапки не валит весь workbook — пропускается с сообщением."""
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = "сломанный"
    ws2 = wb.create_sheet("смета 12.01.01")
    ws2["A1"] = "СМЕТА № 12.01.01"
    ws2["A2"] = ("Длинное наименование объекта строительства которое больше шестидесяти символов точно")
    ws2["A3"] = "№ пп"
    p = tmp_path / "mixed.xlsx"
    wb.save(str(p))

    smetas = parse_workbook(str(p))
    # сломанный лист пропущен, валидный распарсен
    assert len(smetas) == 1
    assert smetas[0].number == "12.01.01"
