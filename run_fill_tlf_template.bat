@echo off
chcp 65001 >nul
REM Fill TLF Template Script Launcher

echo.
echo ================================================================================
echo Fill TLF Template Tool
echo ================================================================================
echo.

REM Get current script directory
set script_dir=%~dp0

REM Resolve Python executable (venv -> py -3.13 -> system python)
set "python_cmd="
if exist "%script_dir%.venv\Scripts\python.exe" (
    "%script_dir%.venv\Scripts\python.exe" --version >nul 2>&1
    if %errorlevel% equ 0 set "python_cmd=%script_dir%.venv\Scripts\python.exe"
)
if not defined python_cmd (
    py -3.13 --version >nul 2>&1
    if %errorlevel% equ 0 set "python_cmd=py -3.13"
)
if not defined python_cmd (
    python --version >nul 2>&1
    if %errorlevel% equ 0 set "python_cmd=python"
)
if not defined python_cmd (
    echo Error: Python not found. Install Python 3.13 and try again.
    pause
    exit /b 1
)

REM Check if dependencies are installed
call %python_cmd% -c "import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Missing dependency openpyxl
    echo Please run: %python_cmd% -m pip install openpyxl
    pause
    exit /b 1
)

REM Run Python script
call %python_cmd% "%script_dir%fill_tlf_template.py"

REM Check return value
if %errorlevel% neq 0 (
    echo.
    echo ================================================================================
    echo Program execution error!
    echo ================================================================================
    pause
    exit /b 1
)

echo.
echo Program completed, press any key to exit...
pause
