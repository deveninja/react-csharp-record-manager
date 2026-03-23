@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PYTHON_CMD="
set "PYTHON_ARGS="

if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    set "PYTHON_CMD=%SCRIPT_DIR%.venv\Scripts\python.exe"
) else (
    where py >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_CMD=py"
        set "PYTHON_ARGS=-3"
    ) else (
        where python >nul 2>nul
        if not errorlevel 1 (
            set "PYTHON_CMD=python"
        )
    )
)

if not defined PYTHON_CMD (
    echo.
    echo ERROR: Could not find a usable Python interpreter.
    echo Install Python 3.9+ from https://www.python.org/downloads/ or create .venv\Scripts\python.exe next to this file.
    pause
    exit /b 1
)

pushd "%SCRIPT_DIR%"
call "%PYTHON_CMD%" %PYTHON_ARGS% main.py
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
    echo.
    echo ERROR: Could not launch the app.
    echo Interpreter used: %PYTHON_CMD% %PYTHON_ARGS%
    pause
    exit /b %EXIT_CODE%
)
