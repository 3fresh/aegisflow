# Cleanup Script - Remove unused files and folders
$workDir = "c:\Users\kplp794\OneDrive - AZCollaboration\Desktop\roooooot\00-工具开发\credit_latest\03_mastertoc"
Set-Location $workDir

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Cleaning up unused files and folders..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Delete __pycache__
if (Test-Path "__pycache__") {
    Remove-Item -Path "__pycache__" -Recurse -Force
    Write-Host "[✓] Deleted __pycache__/" -ForegroundColor Green
} else {
    Write-Host "[!] __pycache__/ not found (already deleted)" -ForegroundColor Yellow
}

# Delete archive
if (Test-Path "archive") {
    Remove-Item -Path "archive" -Recurse -Force
    Write-Host "[✓] Deleted archive/" -ForegroundColor Green
} else {
    Write-Host "[!] archive/ not found (already deleted)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Cleanup completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
