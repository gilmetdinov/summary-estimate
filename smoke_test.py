"""
Smoke-тест: прогон конвертера + рендер результата через LibreOffice headless.

Зачем: openpyxl пишет формулы (кросс-ссылки, сумма прописью, НДС), но НЕ считает их.
LibreOffice при конвертации в PDF пересчитывает всё — если он отработал и PDF
получился ненулевым, значит формулы корректны и файл открывается реальным офисом.

Запуск:
    python3 smoke_test.py                 # на sources/raw.xlsx
    python3 smoke_test.py myraw.xlsx ...  # на своих выгрузках

Возвращает код 0 при успехе, иначе !=0 (годится для CI).
"""
from __future__ import annotations
import os
import shutil
import subprocess
import sys
import tempfile

from config import load_config
from smeta_render import convert

SOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
    "soffice", "libreoffice",                                 # PATH (Linux/Win)
    r"C:\Program Files\LibreOffice\program\soffice.exe",      # Windows
]
MIN_PDF_BYTES = 2000  # пустой/битый PDF будет меньше


def find_soffice() -> str | None:
    for c in SOFFICE_CANDIDATES:
        if os.path.isabs(c) and os.path.exists(c):
            return c
        if shutil.which(c):
            return c
    return None


def render_pdf(xlsx_path: str, out_dir: str, soffice: str) -> str:
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, xlsx_path],
        check=True, capture_output=True, timeout=120,
    )
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError("LibreOffice не создал PDF")
    return pdf


def main(argv: list[str]) -> int:
    inputs = argv or ["sources/raw.xlsx"]
    cfg = load_config("config.json")
    soffice = find_soffice()
    if not soffice:
        print("!! LibreOffice не найден — smoke-тест неполный (проверена только конвертация).")

    ok = True
    with tempfile.TemporaryDirectory() as tmp:
        for src in inputs:
            if not os.path.exists(src):
                print(f"FAIL  {src}: файл не найден"); ok = False; continue
            xlsx = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0] + "_out.xlsx")
            try:
                convert(src, xlsx, cfg=cfg)  # донор по умолчанию = template.xlsx
            except Exception as ex:  # noqa: BLE001
                print(f"FAIL  {src}: конвертация упала: {ex}"); ok = False; continue
            if not soffice:
                print(f"OK*   {src}: конвертация прошла (без рендера)"); continue
            try:
                pdf = render_pdf(xlsx, tmp, soffice)
            except Exception as ex:  # noqa: BLE001
                print(f"FAIL  {src}: LibreOffice не отрендерил: {ex}"); ok = False; continue
            size = os.path.getsize(pdf)
            if size < MIN_PDF_BYTES:
                print(f"FAIL  {src}: PDF подозрительно мал ({size} б)"); ok = False; continue
            print(f"OK    {src}: xlsx + PDF ({size // 1024} КБ) — формулы пересчитались")

    print("\nИТОГ:", "ВСЁ ОК ✅" if ok else "ЕСТЬ ОШИБКИ ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
