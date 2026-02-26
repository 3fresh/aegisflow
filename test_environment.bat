@echo off
chcp 65001 >nul
REM Test Python environment and dependencies

echo ================================================
echo Testing Python Environment and Dependencies
echo ================================================
echo.

REM Resolve Python executable (venv -> py -3.13 -> system python)
set "python_cmd="
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" --version >nul 2>&1
    if %errorlevel% equ 0 set "python_cmd=.venv\Scripts\python.exe"
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

echo Checking Python version...
call %python_cmd% --version
echo.
echo Testing module imports...
call %python_cmd% test_imports.py

echo.
echo ================================================
echo Test Completed
echo ================================================
pause
