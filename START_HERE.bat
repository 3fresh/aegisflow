@echo off
chcp 65001 >nul
cls
echo ===============================================================================
echo TLF TOOLS - Quick Start Menu
echo ===============================================================================
echo.
echo Available Tools:
echo.
echo   1. MOSAIC_CONVERT      - CSV to Excel Tool
echo   2. EXTRACT_PROGRAMS    - Extract SAS Program List
echo   3. FILL_TLF_TEMPLATE   - Fill TLF Template
echo   4. FILL_TLF_STATUS     - Fill TLF Status
echo   5. GENERATE_BATCH_XML  - Generate Batch XML
echo.
echo   0. Test Environment
echo.
echo ===============================================================================
echo.

:menu
choice /C 1234506 /N /M "Select a tool (1-5, 0=test, 6=exit): "

if errorlevel 7 goto end
if errorlevel 6 goto test_env
if errorlevel 5 goto generate_xml
if errorlevel 4 goto fill_status
if errorlevel 3 goto fill_template
if errorlevel 2 goto extract_programs
if errorlevel 1 goto mosaic_convert

:mosaic_convert
echo.
echo Starting MOSAIC_CONVERT...
call run_mosaic_convert.bat
goto menu

:extract_programs
echo.
echo Starting EXTRACT_PROGRAMS...
call run_extract_programs.bat
goto menu

:fill_template
echo.
echo Starting FILL_TLF_TEMPLATE...
call run_fill_tlf_template.bat
goto menu

:fill_status
echo.
echo Starting FILL_TLF_STATUS...
call run_fill_tlf_status.bat
goto menu

:generate_xml
echo.
echo Starting GENERATE_BATCH_XML...
call run_generate_batch_xml.bat
goto menu

:test_env
echo.
echo Running environment test...
call test_environment.bat
pause
goto menu

:end
echo.
echo Thank you for using TLF Tools!
pause
