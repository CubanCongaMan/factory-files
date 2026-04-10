@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"
set "PY_SCRIPT=%SCRIPT_DIR%Televiewer_App_2026_04_10_Rev_01.py"

if not exist "%PY_SCRIPT%" (
    echo Could not find Televiewer_App_2026_04_10_Rev_01.py in this folder.
    echo Expected: %PY_SCRIPT%
    echo.
    pause
    exit /b 1
)

if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    "%SCRIPT_DIR%.venv\Scripts\python.exe" "%PY_SCRIPT%"
) else (
    py "%PY_SCRIPT%" 2>nul || python "%PY_SCRIPT%"
)

echo.
pause
endlocal
