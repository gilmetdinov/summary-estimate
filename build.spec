# PyInstaller spec — автономный .exe веб-конвертера смет ПИР.
#
# ВАЖНО: PyInstaller НЕ умеет кросс-компиляцию. Windows .exe собирается ТОЛЬКО на
# Windows, macOS-бинарь — только на macOS. Для exe под сметчицу запускать на Windows.
#
# Сборка:  pyinstaller build.spec     (результат: dist/summary-estimate[.exe])
#
# В сборку вкладываются template.xlsx (мастер-шаблон оформления) и config.json (дефолты
# формы). Рядом с готовым .exe можно положить свой config.json — он переопределит вложенный.
# Перед сборкой при необходимости перегенерируйте шаблон: python make_template.py
#
# Иконка: положите app.ico в корень проекта — она автоматически вошьётся в exe.
# Нет файла — соберётся со стандартной иконкой (без ошибок).
import os
_ico = os.path.join(SPECPATH, "app.ico")
_ICON = _ico if os.path.exists(_ico) else None   # иконка самого .exe (в проводнике)

# Иконку окна/панели задач tkinter грузит в рантайме сам — поэтому кладём её ВНУТРЬ exe.
_datas = [("template.xlsx", "."), ("config.json", ".")]
for _f in ("app.ico", "app_icon.png"):
    _p = os.path.join(SPECPATH, _f)
    if os.path.exists(_p):
        _datas.append((_p, "."))

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=_datas,
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=['pandas', 'numpy', 'PIL', 'matplotlib'],  # tkinter НУЖЕН — не исключать
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='summary-estimate',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # десктопное окно без отдельной консоли; ошибки -> messagebox + log
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=_ICON,
)
