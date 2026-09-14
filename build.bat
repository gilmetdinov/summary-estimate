@echo off
chcp 65001 >nul
REM ============================================================
REM  Сборка автономного summary-estimate.exe (запускать на Windows).
REM  PyInstaller кросс-компиляцию не умеет — exe собирается только на винде.
REM  Нужен установленный Python 3.10+ (с офсайта python.org — в нём есть всё нужное).
REM ============================================================
setlocal

REM venv создаём только если его ещё нет (иначе при запуске изнутри активного .venv
REM Windows не даст перезаписать python.exe, будет ошибка "Unable to copy venvlauncher.exe").
if not exist ".venv\Scripts\python.exe" (
    py -3 -m venv .venv 2>nul || python -m venv .venv || goto :err
)
call .venv\Scripts\activate.bat || goto :err
python -m pip install --upgrade pip
pip install -r requirements-build.txt || goto :err
pyinstaller --noconfirm --clean build.spec || goto :err

echo.
echo ============================================================
echo  Готово:  dist\summary-estimate.exe
echo  Отдавайте сметчице ОДИН этот файл. Двойной клик - откроется окно программы.
echo ============================================================
goto :end

:err
echo.
echo  ОШИБКА сборки. Проверьте, что установлен Python и есть интернет для pip.

:end
pause
endlocal
