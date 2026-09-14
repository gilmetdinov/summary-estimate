"""
Конфиг реквизитов и параметров прогона (вынесен из хардкода в smeta_render.py).

Подписанты (проектная организация / заказчик), ставка НДС и опциональные строки
сводной (изыскания НЗ_ИГДИ и госэкспертиза — они НЕ приходят из выгрузки ГС,
вставляются вручную). Значения по умолчанию совпадают с прежним хардкодом, чтобы
прогон без config.json давал тот же результат, что и раньше.
"""
from __future__ import annotations
import dataclasses
import json
import os
from dataclasses import dataclass


def _num(v):
    """'1 234,56' | 1234.56 | '' -> float | None (терпим к запятой/пробелам)."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _none_or(v, default):
    """v если не None, иначе default (в отличие от `or` не съедает 0.0)."""
    return v if v is not None else default


@dataclass
class Placeholder:
    """Опциональная строка сводной, вставляемая вручную (изыскания/госэкспертиза)."""
    name: str = ""          # наименование (колонка B)
    osn: str = ""           # обоснование (колонка D)
    value: float | None = None  # стоимость, тыс. руб.
    column: str = "izysk"   # 'izysk' -> колонка инж. изысканий (E), 'proj' -> проектных работ (F)


@dataclass
class Config:
    # --- подписант проектной организации ---
    org_short: str = "ООО «Проектная организация»"
    org_dir_title1: str = "Заместитель генерального директора -"
    org_dir_title2: str = "директор по проектированию"
    org_dir_name: str = "И.И. Иванов"
    gip_name: str = "П.П. Петров"
    # --- подписант заказчика ---
    cust_dir_title: str = "Директор"
    cust_dir_title2: str = "Заказчик"
    cust_dir_name: str = "С.С. Сидоров"
    # --- ставка НДС (доля: 0.22 == 22%) ---
    vat_rate: float = 0.22
    # --- опциональные строки сводной (вставляются вручную) ---
    izyskaniya: Placeholder | None = None
    gosexpertiza: Placeholder | None = None
    # --- ед. стоимость для колонок O/P сводной (ввод из формы, пусто = не заполнено) ---
    unit_cost_proj: float | None = None
    unit_cost_izysk: float | None = None
    unit_cost_gosexp: float | None = None

    @property
    def vat_pct(self) -> str:
        """0.22 -> '22', 0.205 -> '20.5' (для подписи 'НДС N%')."""
        return f"{self.vat_rate * 100:g}"


def _placeholder_from(v, default_column: str) -> Placeholder | None:
    if not v:
        return None
    if isinstance(v, dict):
        return Placeholder(
            name=str(v.get("name", "")),
            osn=str(v.get("osn", "")),
            value=_num(v.get("value")),
            column=str(v.get("column", default_column)),
        )
    return None


def from_dict(data: dict) -> Config:
    data = dict(data or {})
    data.pop("_comment", None)
    data["izyskaniya"] = _placeholder_from(data.get("izyskaniya"), "izysk")
    data["gosexpertiza"] = _placeholder_from(data.get("gosexpertiza"), "proj")
    if "vat_rate" in data and data["vat_rate"] is not None:
        data["vat_rate"] = float(data["vat_rate"])
    allowed = {f.name for f in dataclasses.fields(Config)}
    data = {k: v for k, v in data.items() if k in allowed}
    return Config(**data)


def load_config(path: str | None = "config.json") -> Config:
    """Читает config.json; если файла нет — дефолты (= прежний хардкод)."""
    if not path or not os.path.exists(path):
        return Config()
    with open(path, encoding="utf-8") as f:
        return from_dict(json.load(f))


def to_dict(cfg: Config) -> dict:
    d = dataclasses.asdict(cfg)
    for k in ("izyskaniya", "gosexpertiza"):
        if d.get(k) is None:
            d[k] = None
    return d


def save_config(cfg: Config, path: str = "config.json") -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(to_dict(cfg), f, ensure_ascii=False, indent=2)


def config_from_strings(values: dict, base: Config | None = None) -> Config:
    """Строит Config из «сырых» строковых полей формы (GUI/веб).

    Пустое поле -> значение из base (дефолт). Чекбоксы изысканий/госэкспертизы:
    любое истинное значение ('on', True, 1) включает соответствующую строку.
    Вынесено сюда (без зависимостей от tkinter), чтобы покрывалось тестами headless.
    """
    base = base or Config()

    def g(key, default):
        v = values.get(key)
        v = "" if v is None else str(v).strip()
        return v or default

    cfg = Config(
        org_short=g("org_short", base.org_short),
        org_dir_title1=g("org_dir_title1", base.org_dir_title1),
        org_dir_title2=g("org_dir_title2", base.org_dir_title2),
        org_dir_name=g("org_dir_name", base.org_dir_name),
        gip_name=g("gip_name", base.gip_name),
        cust_dir_title=g("cust_dir_title", base.cust_dir_title),
        cust_dir_title2=g("cust_dir_title2", base.cust_dir_title2),
        cust_dir_name=g("cust_dir_name", base.cust_dir_name),
    )
    vp = _num(values.get("vat_pct"))
    cfg.vat_rate = vp / 100 if vp is not None else base.vat_rate
    if values.get("izysk_on"):
        cfg.izyskaniya = Placeholder(
            name=str(values.get("izysk_name", "")).strip(),
            osn=str(values.get("izysk_osn", "")).strip(),
            value=_num(values.get("izysk_value")), column="izysk")
    if values.get("gosexp_on"):
        cfg.gosexpertiza = Placeholder(
            name=str(values.get("gosexp_name", "")).strip(),
            osn=str(values.get("gosexp_osn", "")).strip(),
            value=_num(values.get("gosexp_value")), column="proj")
    cfg.unit_cost_proj = _none_or(_num(values.get("unit_cost_proj")), base.unit_cost_proj)
    cfg.unit_cost_izysk = _none_or(_num(values.get("unit_cost_izysk")), base.unit_cost_izysk)
    cfg.unit_cost_gosexp = _none_or(_num(values.get("unit_cost_gosexp")), base.unit_cost_gosexp)
    return cfg
