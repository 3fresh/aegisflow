# Fix BAT file encoding - Remove BOM to fix @echo off issue
Write-Host "Starting BAT file encoding fix..." -ForegroundColor Green
Write-Host ""

$batFiles = @(
    "START_HERE.bat",
    "run_extract_programs.bat",
    "run_mosaic_convert.bat",
    "run_generate_batch_xml.bat",
    "run_fill_tlf_status.bat",
    "run_fill_tlf_template.bat",
    "test_environment.bat"
)

$basePath = $PSScriptRoot

foreach ($file in $batFiles) {
    $path = Join-Path $basePath $file
    
    if (Test-Path $path) {
        try {
            # Read file content
            $content = Get-Content $path -Raw -Encoding UTF8
            
            # Write back with UTF-8 without BOM
            $utf8NoBom = New-Object System.Text.UTF8Encoding $false
            [System.IO.File]::WriteAllText($path, $content, $utf8NoBom)
            
            Write-Host "  [OK] $file" -ForegroundColor Cyan
        }
        catch {
            Write-Host "  [FAIL] $file - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "  [SKIP] $file - File not found" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "Encoding fix completed!" -ForegroundColor Green
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
