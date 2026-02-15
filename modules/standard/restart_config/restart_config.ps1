# ========================================
# Restart with AutoRun
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Restart with AutoRun" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Resolve Fabriq.bat absolute path
# ========================================
$fabriqRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$fabriqBat = Join-Path $fabriqRoot "Fabriq.bat"

if (-not (Test-Path $fabriqBat)) {
    Write-Host "[ERROR] Fabriq.bat not found: $fabriqBat" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Fabriq.bat not found: $fabriqBat")
}

Write-Host "[INFO] Fabriq.bat: $fabriqBat" -ForegroundColor Cyan
Write-Host ""

# ========================================
# RunOnce Registry Settings
# ========================================
$runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
$runOnceName = "FabriqAutoStart"
$runOnceValue = "cmd /c `"$fabriqBat`""

# Check current RunOnce state
$existingValue = Get-ItemProperty -Path $runOncePath -Name $runOnceName -ErrorAction SilentlyContinue

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following RunOnce entry will be registered" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Path:  $runOncePath"
Write-Host "  Name:  $runOnceName"
Write-Host "  Value: $runOnceValue"
Write-Host ""

if ($existingValue) {
    Write-Host "  [Current] RunOnce entry already exists: $($existingValue.$runOnceName)" -ForegroundColor Gray
} else {
    Write-Host "  [Change] RunOnce entry will be created" -ForegroundColor White
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Register RunOnce and restart the computer?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Register RunOnce
# ========================================
try {
    if (-not (Test-Path $runOncePath)) {
        New-Item -Path $runOncePath -Force | Out-Null
    }

    $existing = Get-ItemProperty -Path $runOncePath -Name $runOnceName -ErrorAction SilentlyContinue
    if ($existing) {
        Set-ItemProperty -Path $runOncePath -Name $runOnceName -Value $runOnceValue -Type String -Force -ErrorAction Stop
    } else {
        New-ItemProperty -Path $runOncePath -Name $runOnceName -Value $runOnceValue -PropertyType String -Force -ErrorAction Stop | Out-Null
    }

    Write-Host "[SUCCESS] RunOnce registered: $runOnceName" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Failed to register RunOnce: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to register RunOnce: $_")
}

# ========================================
# Countdown and Restart
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The computer will restart in 10 seconds..." -ForegroundColor Yellow
Write-Host "Press Ctrl+C to abort" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

for ($i = 10; $i -ge 1; $i--) {
    Write-Host "`r  Restarting in $i seconds... " -NoNewline -ForegroundColor Yellow
    Start-Sleep -Seconds 1
}
Write-Host ""
Write-Host ""

Write-Host "[INFO] Restarting computer..." -ForegroundColor Cyan

# Write result before restart (execution history will be saved by the framework)
$result = New-ModuleResult -Status "Success" -Message "RunOnce registered, restarting computer"

Restart-Computer -Force

return $result
