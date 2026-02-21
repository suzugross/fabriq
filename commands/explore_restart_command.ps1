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

Write-Host "[INFO] Waiting for Windows to restart Explorer automatically..." -ForegroundColor Cyan

$maxWait  = 15
$interval = 1
$elapsed  = 0
$restarted = $false

while ($elapsed -lt $maxWait) {
    Start-Sleep -Seconds $interval
    $elapsed += $interval
    if (@(Get-Process -Name explorer -ErrorAction SilentlyContinue).Count -gt 0) {
        $restarted = $true
        break
    }
}

if ($restarted) {
    Write-Host "[SUCCESS] Explorer restarted (${elapsed}s)" -ForegroundColor Green
} else {
    Write-Host "[WARNING] Explorer did not restart within ${maxWait}s. Please check manually." -ForegroundColor Yellow
}

Write-Host ""