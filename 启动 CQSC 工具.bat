@echo off
REM ============================================================
REM  CQSC Tools v2.0 launcher
REM  - Lock to project dir via %~dp0 (works no matter how launched)
REM  - Use the project's own .venv (Python 3.12 + pywin32)
REM  - Clear PYTHONPATH/PYTHONHOME so external env can't leak in
REM  - Use chcp 65001 + start "" so the bat window disappears
REM ============================================================

setlocal
chcp 65001 >nul
cd /d "%~dp0"

set PYTHONPATH=
set PYTHONHOME=

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo [ERROR] venv not found: %VENV_PY%
    echo Create it with:  python -m venv .venv
    echo Then install:    .venv\Scripts\python.exe -m pip install pywin32
    pause
    exit /b 1
)

start "" /b "%VENV_PY%" "%~dp0main.py"
endlocal
exit /b 0