@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PY_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe"
set "APP_SCRIPT=%SCRIPT_DIR%Photo_Selector_V26_Rev_15.py"

if not exist "%PY_EXE%" (
    echo ERROR: Python executable not found at "%PY_EXE%"
    pause
    exit /b 1
)

if not exist "%APP_SCRIPT%" (
    echo ERROR: App script not found at "%APP_SCRIPT%"
    pause
    exit /b 1
)

echo Launching Photo Selector Rev 15...
"%PY_EXE%" "%APP_SCRIPT%"

if errorlevel 1 (
    echo.
    echo App exited with an error.
)

pause
endlocal
