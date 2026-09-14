#!/usr/bin/env bash
# Сборка автономного бинаря на macOS/Linux (для локальной проверки spec и вложения
# ресурсов). Готовый файл — dist/summary-estimate (НЕ .exe; Windows-exe собирается
# только на Windows через build.bat — кросс-компиляции у PyInstaller нет).
set -e
cd "$(dirname "$0")"

python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements-build.txt
pyinstaller build.spec

echo
echo "Готово: dist/summary-estimate"
echo "Запуск: ./dist/summary-estimate  — откроется окно программы."
