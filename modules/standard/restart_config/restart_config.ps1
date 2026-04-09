# ========================================
# Restart with AutoRun
# ========================================

Write-Host ""
Show-Separator
Write-Host "Restart with AutoRun" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Resolve Fabriq entry point (prefer .exe, fallback .bat)
# ========================================
$fabriqRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$fabriqExe = Join-Path $fabriqRoot "Fabriq.exe"
$fabriqBat = Join-Path $fabriqRoot "Fabriq.bat"

if (Test-Path $fabriqExe) {
    $runOnceValue = "`"$fabriqExe`""
    $entryPoint = $fabriqExe
} elseif (Test-Path $fabriqBat) {
    $runOnceValue = "cmd /c `"$fabriqBat`""
    $entryPoint = $fabriqBat
} else {
    Show-Error "Fabriq entry point not found (neither .exe nor .bat)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Fabriq entry point not found")
}

Show-Info "Entry point: $entryPoint"
Write-Host ""

# ========================================
# RunOnce Registry Settings
# ========================================
$runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
$runOnceName = "FabriqAutoStart"

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
$cancelResult = Confirm-ModuleExecution -Message "Register RunOnce and restart the computer?"
if ($null -ne $cancelResult) { return $cancelResult }

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

    Show-Success "RunOnce registered: $runOnceName"
    Write-Host ""
}
catch {
    Show-Error "Failed to register RunOnce: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to register RunOnce: $_")
}

# ========================================
# Countdown and Restart
# ========================================
# Write result before restart (execution history will be saved by the framework)
$result = New-ModuleResult -Status "Success" -Message "RunOnce registered, restarting computer"

Invoke-CountdownRestart -Seconds 10

return $result
