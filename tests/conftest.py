"""Общие фикстуры и хелперы для тестов.

Синтетическая «выгрузка ГС» строится прямо в тесте через openpyxl — тесты
самодостаточны и не зависят от локальных sources/ (под .gitignore).
"""
from __future__ import annotations

import openpyxl
import pytest


def make_sheet(rows: list[tuple]) -> "openpyxl.worksheet.worksheet.Worksheet":
    """Строит лист из списка строк. Каждая строка — кортеж значений A..E."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "смета 12.01.01"
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            if val is not None:
                ws.cell(row=r, column=c).value = val
    return ws


def make_workbook(path: str, rows: list[tuple]) -> str:
    """Сохраняет синтетическую выгрузку в xlsx и возвращает путь."""
    ws = make_sheet(rows)
    ws.parent.save(path)
    return path


@pytest.fixture
def raw_xlsx(tmp_path):
    """Файл выгрузки с одной сметой (шапка + раздел + 2 позиции + итоги)."""
    p = tmp_path / "raw.xlsx"
    rows = [
        ("СМЕТА № 12.01.01", None, None, None, None),
        ("Строительство объекта капитального строительства — пример очень длинного наименования стройки", None, None, None, None),
        ("на проектные работы", None, None, None, None),
        ("Составлена в уровне цен на 1 квартал 2025 г.", None, None, None, None),
        ("Заказчик", None, None, None, None),
        (None, "ООО «Заказчик»", None, None, None),
        ("Проектная организация", None, None, None, None),
        (None, "ООО «Проектная организация»", None, None, None),
        ("№ пп", None, None, None, None),
        ("1", None, None, None, None),
        ("Раздел 1. Проектные работы", None, None, None, None),
        ("1", "Позиция один", "МНЗ п.1", "расчёт 1", 100.5),
        (None, None, "коэффициент 0,9", None, 12.5),
        ("2", "Позиция два", "МНЗ п.2", "расчёт 2", 50.0),
        (None, "Итого Поз. 1-2", None, None, 163.0),
        (None, "НДС 22%", None, None, 35.86),
        (None, "Всего по смете", None, None, 198.86),
    ]
    return make_workbook(str(p), rows)
