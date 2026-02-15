# ========================================
# Explorer Restart Script
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Restart Explorer" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[INFO] Restarting Explorer" -ForegroundColor Cyan
Write-Host "[INFO] The taskbar and desktop will temporarily disappear. This is normal." -ForegroundColor Yellow
Write-Host ""

if (-not (Confirm-Execution -Message "Do you want to execute?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return
}

Write-Host ""
Write-Host "[INFO] Stopping explorer.exe..." -ForegroundColor Cyan

try {
    Stop-Process -Name explorer -Force -ErrorAction Stop
    Write-Host "[SUCCESS] Stopped explorer.exe" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to stop explorer.exe: $_" -ForegroundColor Red
    Write-Host ""
    return
}

Start-Sleep -Seconds 1

Write-Host "[INFO] Starting explorer.exe..." -ForegroundColor Cyan

try {
    Start-Process explorer.exe -ErrorAction Stop
    Write-Host "[SUCCESS] Started explorer.exe" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to start explorer.exe: $_" -ForegroundColor Red
    Write-Host "[INFO] Please start explorer.exe manually" -ForegroundColor Yellow
}

Write-Host ""