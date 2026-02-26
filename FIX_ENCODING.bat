@echo off
echo.
echo ========================================
echo BAT File Encoding Fix
echo ========================================
echo.
echo This will remove BOM from BAT files to fix @echo off issue
echo.
pause
echo.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0fix_bat_encoding.ps1"
