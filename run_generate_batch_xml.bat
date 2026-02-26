@echo off
chcp 65001 >nul
REM Generate Batch List XML from Excel

echo ================================================
echo Generate Batch List XML Tool
echo ================================================
echo.

REM Resolve Python executable (venv -> py -3.13 -> system python)
set "python_cmd="
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "python_cmd=.venv\Scripts\python.exe"
        echo Using Python from virtual environment...
    )
)
if not defined python_cmd (
    py -3.13 --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "python_cmd=py -3.13"
        echo Using Python 3.13 via launcher...
    )
)
if not defined python_cmd (
    python --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "python_cmd=python"
        echo Using system Python...
    )
)
if not defined python_cmd (
    echo Error: Python not found. Install Python 3.13 and try again.
    pause
    exit /b 1
)

call %python_cmd% generate_batch_xml.py

if %errorlevel% neq 0 (
    echo.
    echo ================================================
    echo Error: Program execution failed ^(error code: %errorlevel%^)
    echo ================================================
    echo.
    echo Possible causes:
    echo 1. Missing required Python packages ^(pandas, openpyxl^)
    echo 2. Incompatible Python version
    echo 3. tkinter not installed
    echo.
    echo Try running: pip install pandas openpyxl
    echo.
)

echo.
echo ================================================
pause
