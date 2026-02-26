@echo off
chcp 65001 >nul
echo ================================================================
echo MOSAIC_CONVERT v3.1 - CSV to Excel Tool (seq transpose)
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

echo [1] Checking CSV structure and performing seq transpose...
echo.
call %python_cmd% mosaic_convert.py

if errorlevel 1 (
    echo.
    echo ERROR: Conversion failed - Check for duplicate tocnumbers in CSV
    pause
    exit /b 1
)

echo.
echo.
echo [2] Validating output file (249 rows expected)...
echo.

REM Call Python script directly, let Python read .last_output.txt (avoid bat passing Chinese path)
if exist ".last_output.txt" (
    call %python_cmd% validate_output.py
) else (
    echo WARN: Output file path not found, skipping validation
)

echo.
echo ================================================================
echo Completed!
echo ================================================================
echo Note: If program+suffix duplicate warnings appear, update MOSAIC macro
echo ================================================================
pause
