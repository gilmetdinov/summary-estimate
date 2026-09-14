"""
Мастер-шаблон-донор оформления.

Рендерер не генерирует стили вручную: он копирует оформление (шрифт, рамки, заливку,
выравнивание, формат чисел) из выделенного файла-донора ``template.xlsx``. Здесь — класс
доступа к донору и имена листов-доноров.
"""
from __future__ import annotations
import re

import openpyxl

# имена листов-доноров в мастер-шаблоне template.xlsx (см. make_template.py)
SMETA_DONOR = "ШАБЛОН_смета"
SVOD_DONOR = "ШАБЛОН_сводная"
DEFAULT_DONOR = "template.xlsx"


class Donor:
    def __init__(self, path: str):
        self._w = openpyxl.load_workbook(path)
        self._wf = openpyxl.load_workbook(path, data_only=False)
        self.sm = self._w[SMETA_DONOR]
        self.sv = self._w[SVOD_DONOR]
        # формулы «сумма прописью» (шаблоны)
        self.propis_smeta = str(self._wf[SMETA_DONOR]["A55"].value)
        self.propis_svod = str(self._wf[SVOD_DONOR]["C30"].value)

    def propis_for_smeta(self, addr: str) -> str:
        return re.sub(r"\bE53\b", addr,
                      self.propis_smeta.replace("с НДС", "без НДС"))

    def propis_for_svod(self, addr: str) -> str:
        return re.sub(r"\bG28\b", addr, self.propis_svod)
