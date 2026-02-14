# ========================================
# Group Policy Update Script
# ========================================

Write-Host "Forcing Group Policy update..." -ForegroundColor Cyan
Write-Host ""

try {
    gpupdate /force
    Write-Host ""
    Write-Host "[SUCCESS] Group Policy update completed" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "[ERROR] Error occurred during Group Policy update: $_" -ForegroundColor Red
}

Write-Host ""
pause