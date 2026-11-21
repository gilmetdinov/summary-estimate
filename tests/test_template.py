"""Мастер-шаблон template.xlsx не должен нести данные/метаданные исходной сметы.

Ловит регресс, когда make_template.py оставлял все формулы (ссылки на листы исходной
книги, захардкоженные суммы) и свойства документа (автор, кем изменён).
"""
from __future__ import annotations

import os
import zipfile

import openpyxl
import pytest

import make_template
from render.donor import DEFAULT_DONOR, SMETA_DONOR, SVOD_DONOR

pytestmark = pytest.mark.skipif(not os.path.exists(DEFAULT_DONOR), reason="нет template.xlsx")


def _assert_clean(path: str):
    wb = openpyxl.load_workbook(path)
    assert set(wb.sheetnames) == {SVOD_DONOR, SMETA_DONOR}

    # свойства документа — нейтральные
    p = wb.properties
    assert p.creator == make_template.NEUTRAL_CREATOR
    assert not p.lastModifiedBy
    assert not p.lastPrinted
    for attr in ("title", "subject", "description", "keywords", "category"):
        assert not getattr(p, attr), attr

    # только формулы «сумма прописью», без ссылок на другие листы; никаких данных
    for ws in wb.worksheets:
        keep = make_template.KEEP_FORMULAS[ws.title]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                assert cell.coordinate in keep, f"{ws.title}!{cell.coordinate} = {str(v)[:40]!r}"
                assert isinstance(v, str) and v.startswith("=")
                assert "!" not in v, f"кросс-ссылка в {ws.title}!{cell.coordinate}"
                assert "[" not in v, f"ссылка на внешнюю книгу в {ws.title}!{cell.coordinate}"
                assert cell.comment is None
        assert not ws.defined_names
        assert not ws.auto_filter.ref
    assert not wb.defined_names
    assert not wb._external_links

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        assert not [n for n in names if "externalLink" in n or "comments" in n]
        core = z.read("docProps/core.xml").decode()
        assert "lastModifiedBy" not in core and "lastPrinted" not in core


def test_template_is_clean():
    _assert_clean(DEFAULT_DONOR)


def test_template_keeps_formulas_renderer_needs():
    from render.donor import Donor
    d = Donor(DEFAULT_DONOR)
    assert "E53" in d.propis_smeta and d.propis_smeta.startswith("=")
    assert "G28" in d.propis_svod and d.propis_svod.startswith("=")


def test_make_template_scrubs_dirty_source(tmp_path):
    """make_template.build на «грязном» эталоне даёт чистый шаблон."""
    src = tmp_path / "example.xlsx"
    wb = openpyxl.Workbook()
    svod = wb.active
    svod.title = "св.смета_77"
    sm = wb.create_sheet("смета 77.01.01_тест")
    wb.create_sheet("смета 77.02.01_лишняя")
    wb.create_sheet("прочее")
    sm["A55"] = '=CONCATENATE("Всего по смете с НДС: ",E53)'
    sm["E1"] = '=CONCATENATE("Приложение №3.",F1)'
    sm["B12"] = "ООО «Реальный заказчик»"
    sm["E51"] = 123.45
    svod["C30"] = "=G28*1"
    svod["F20"] = "='смета 77.01.01_тест'!E51"
    svod["F22"] = "='смета 77.02.01_лишняя'!E41"
    svod["H26"] = "=98765.43/1000"
    svod["C11"] = "Объект с реальным названием"
    svod.auto_filter.ref = "A18:G28"
    from openpyxl.comments import Comment
    svod["A1"].comment = Comment("личная пометка", "Автор")
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names["secret"] = DefinedName("secret", attr_text="'прочее'!$A$1")
    wb.properties.creator = "Иван Реальный"
    wb.properties.lastModifiedBy = "WORKPC"
    wb.properties.title = "Смета по объекту"
    wb.save(src)

    dst = tmp_path / "template.xlsx"
    make_template.build(str(src), str(dst))
    _assert_clean(str(dst))

    # повторная очистка уже готового шаблона (путь `make_template.py template.xlsx`)
    again = tmp_path / "again.xlsx"
    make_template.build(str(dst), str(again))
    _assert_clean(str(again))
