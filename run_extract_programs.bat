@echo off
chcp 65001 >nul
echo ================================================================
echo EXTRACT_PROGRAMS v1.0 - Extract SAS Program List Tool
echo ================================================================
echo.
echo Features:
echo   1. Extract PROGRAM list from MOSAIC_CONVERT Excel output
echo   2. Count tocnumbers for each program
echo   3. Maintain original order from Excel
echo   4. Comments centralized at file top
echo   5. Generate SAS script format text file
echo.
echo ================================================================
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
    echo ERROR: Python not found. Install Python 3.13 and try again.
    pause
    exit /b 1
)

echo Starting program...
echo.
call %python_cmd% extract_programs.py

if errorlevel 1 (
    echo.
    echo ERROR: Program execution failed
    pause
    exit /b 1
)

echo.
echo ================================================================
echo Program completed successfully!
echo ================================================================
pause
