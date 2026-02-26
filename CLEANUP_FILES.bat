@echo off
chcp 65001 >nul
REM Cleanup unused files and folders

echo ===============================================================================
echo Cleaning up unused files and folders...
echo ===============================================================================
echo.

if exist "__pycache__" (
    echo Deleting __pycache__/...
    rd /s /q "__pycache__"
    if exist "__pycache__" (
        echo [ERROR] Failed to delete __pycache__/
    ) else (
        echo [OK] Deleted __pycache__/
    )
) else (
    echo [INFO] __pycache__/ not found
)

echo.

if exist "archive" (
    echo Deleting archive/...
    rd /s /q "archive"
    if exist "archive" (
        echo [ERROR] Failed to delete archive/
    ) else (
        echo [OK] Deleted archive/
    )
) else (
    echo [INFO] archive/ not found
)

echo.
echo ===============================================================================
echo Cleanup completed!
echo ===============================================================================
echo.

REM Clean up temporary scripts
if exist "cleanup.py" (
    del /q "cleanup.py"
    echo [OK] Removed temporary cleanup.py
)
if exist "cleanup_files.ps1" (
    del /q "cleanup_files.ps1"
    echo [OK] Removed temporary cleanup_files.ps1
)

echo.
pause
