# ========================================
# Edge Profile Restore (with Auto Kill)
# ========================================
# Restores Edge User Data from backup directory using robocopy mirror.
# Automatically terminates Edge processes before restore.
# WARNING: This overwrites current Edge settings completely.
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Edge Profile Restore" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- Paths ---
$backupDir = Join-Path $PSScriptRoot "backup"
$targetDir = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"

Write-Host "  Source:      $backupDir" -ForegroundColor White
Write-Host "  Destination: $targetDir" -ForegroundColor White
Write-Host "  Mode:        MIRROR (Overwrite)" -ForegroundColor Gray
Write-Host ""

# --- Backup existence check ---
if (-not (Test-Path $backupDir)) {
    Write-Host "[ERROR] Backup directory not found: $backupDir" -ForegroundColor Red
    Write-Host "[INFO] Run 'Edge Profile Backup' first to create a backup." -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Backup directory not found")
}

# Verify backup has content
$backupFiles = @(Get-ChildItem $backupDir -Recurse -File -ErrorAction SilentlyContinue)
if ($backupFiles.Count -eq 0) {
    Write-Host "[ERROR] Backup directory is empty: $backupDir" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Backup directory is empty")
}

$backupSize = ($backupFiles | Measure-Object -Property Length -Sum).Sum
$sizeStr = if ($backupSize -gt 1GB) {
    "{0:N1} GB" -f ($backupSize / 1GB)
} elseif ($backupSize -gt 1MB) {
    "{0:N0} MB" -f ($backupSize / 1MB)
} else {
    "{0:N0} KB" -f ($backupSize / 1KB)
}

Write-Host "  Backup size: $sizeStr ($($backupFiles.Count) files)" -ForegroundColor White
Write-Host ""

# --- Edge process check & kill ---
$edgeProcesses = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue)
if ($edgeProcesses.Count -gt 0) {
    Write-Host "[WARNING] Edge is running ($($edgeProcesses.Count) processes)" -ForegroundColor Yellow
    Write-Host "[INFO] Edge must be closed before restore" -ForegroundColor Yellow
    Write-Host ""

    if (-not (Confirm-Execution -Message "Terminate Edge and proceed with restore?")) {
        Write-Host ""
        Write-Host "[INFO] Canceled" -ForegroundColor Yellow
        Write-Host ""
        return (New-ModuleResult -Status "Cancelled" -Message "User canceled (Edge running)")
    }

    Write-Host ""
    Write-Host "[INFO] Terminating Edge processes..." -ForegroundColor Cyan

    try {
        Stop-Process -Name "msedge" -Force -ErrorAction Stop
        Start-Sleep -Seconds 2

        $remaining = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue)
        if ($remaining.Count -gt 0) {
            Write-Host "[ERROR] Edge processes still running after termination attempt" -ForegroundColor Red
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Failed to terminate Edge")
        }

        Write-Host "[SUCCESS] Edge terminated" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Failed to terminate Edge: $_" -ForegroundColor Red
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to terminate Edge: $_")
    }

    Write-Host ""
}
else {
    Write-Host "[INFO] Edge is not running" -ForegroundColor Gray
    Write-Host ""
}

# --- Confirm restore (with overwrite warning) ---
Write-Host "[WARNING] This will OVERWRITE current Edge settings completely." -ForegroundColor Red
Write-Host ""
if (-not (Confirm-Execution -Message "Restore Edge profile from backup?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# --- Execute Robocopy ---
Write-Host "[INFO] Restore in progress (this may take a few minutes)..." -ForegroundColor Cyan
Write-Host ""

& robocopy.exe "$backupDir" "$targetDir" /MIR /XJ /MT /R:1 /W:1 /NFL /NDL
$exitCode = $LASTEXITCODE

Write-Host ""

# --- Result ---
if ($exitCode -lt 8) {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Restore Results" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Status:   Success (robocopy exit: $exitCode)" -ForegroundColor Green
    Write-Host "  Restored: $sizeStr" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    return (New-ModuleResult -Status "Success" -Message "Restore completed ($sizeStr)")
}
else {
    Write-Host "[ERROR] Restore failed (robocopy exit code: $exitCode)" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Robocopy failed (exit: $exitCode)")
}
