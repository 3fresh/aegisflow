@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ============================================================
echo   AegisFlow EXE Builder
echo ============================================================
echo.

REM ── Locate Python ──────────────────────────────────────────────────────
set "python_cmd="
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" --version >nul 2>&1
    if !errorlevel! equ 0 set "python_cmd=.venv\Scripts\python.exe"
)
if not defined python_cmd (
    py -3.13 --version >nul 2>&1
    if !errorlevel! equ 0 set "python_cmd=py -3.13"
)
if not defined python_cmd (
    python --version >nul 2>&1
    if !errorlevel! equ 0 set "python_cmd=python"
)
if not defined python_cmd (
    echo ERROR: No Python interpreter found.
    echo Please install Python 3.13 or activate the .venv virtual environment.
    pause & exit /b 1
)
echo [1] Using Python: %python_cmd%
echo.

REM ── Install build dependencies ──────────────────────────────────────────
echo [2] Installing / updating required packages ...
call "%python_cmd%" -m pip install -q --upgrade pip
call "%python_cmd%" -m pip install -q customtkinter pyinstaller pandas openpyxl pillow
if !errorlevel! neq 0 (
    echo ERROR: Package installation failed.
    pause & exit /b 1
)
echo     Done.
echo.

REM ── Clean previous build artefacts ─────────────────────────────────────
echo [3] Cleaning previous build artefacts ...
if exist "dist\AegisFlow.exe" del /f /q "dist\AegisFlow.exe"
if exist "build" rmdir /s /q build
echo     Done.
echo.

REM ── Run PyInstaller ─────────────────────────────────────────────────────
echo [4] Building EXE with PyInstaller ...
echo     (this may take 1-3 minutes on first run)
echo.
call "%python_cmd%" -m PyInstaller aegisflow.spec --noconfirm

if !errorlevel! neq 0 (
    echo.
    echo ERROR: PyInstaller build failed — see output above.
    pause & exit /b 1
)

REM ── Check output ────────────────────────────────────────────────────────
if exist "dist\AegisFlow.exe" (
    echo.
    echo ============================================================
    echo   BUILD SUCCESSFUL
    echo ============================================================
    echo.
    echo   Output: dist\AegisFlow.exe
    echo.
    echo   You can now:
    echo     - Double-click dist\AegisFlow.exe to launch the app
    echo     - Share dist\AegisFlow.exe with users who don't have Python
    echo.
    for %%F in ("dist\AegisFlow.exe") do echo   File size: %%~zF bytes
) else (
    echo.
    echo ERROR: dist\AegisFlow.exe not found after build.
)

echo.
pause
