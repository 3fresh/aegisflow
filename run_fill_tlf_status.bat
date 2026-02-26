@echo off
chcp 65001 >nul
REM ============================================================
REM TLF Status Filler - Batch Launcher Script
REM Function: Merge Comparison Status from TFL Status to People Management
REM ============================================================

echo.
echo ========================================
echo TLF Status Filler
echo ========================================
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

call %python_cmd% fill_tlf_status.py

REM Check exit code
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo Script executed successfully!
    echo ========================================
) else (
    echo.
    echo ========================================
    echo Script execution failed, check error messages
    echo ========================================
)

echo.
pause
