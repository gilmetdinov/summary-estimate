"""Единственный источник версии программы.

Формат — календарный (CalVer): дата релиза. Сметчица отличает сборки exe
по версии в заголовке окна и в свойствах файла (см. build.spec).
Бамп версии = поменять строку, поставить git-тег v<версия>, запушить тег —
GitHub Actions соберёт exe и выложит в Releases.
"""
from __future__ import annotations

__version__ = "2026.10.15"


def version_tuple() -> tuple[int, int, int, int]:
    """'2026.10.15' -> (2026, 10, 15, 0) — для VSVersionInfo в build.spec."""
    parts = [int(p) for p in __version__.split(".")]
    parts += [0] * (4 - len(parts))
    return tuple(parts[:4])  # type: ignore[return-value]
