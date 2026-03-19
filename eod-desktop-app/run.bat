@echo off
python "%~dp0main.py"
if errorlevel 1 (
    echo.
    echo ERROR: Could not launch the app.
    echo Make sure Python 3.9+ is installed and available in your PATH.
    echo Download Python from https://www.python.org/downloads/
    pause
)
