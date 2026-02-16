# ========================================
# Edge Profile Restore (with Auto Kill)
# ========================================
# Restores Edge User Data from backup directory using robocopy mirror.
# Automatically terminates Edge processes before restore.
# WARNING: This overwrites current Edge settings completely.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Edge Profile Restore" -ForegroundColor Cyan
Show-Separator
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
    Show-Error "Backup directory not found: $backupDir"
    Show-Info "Run 'Edge Profile Backup' first to create a backup."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Backup directory not found")
}

# Verify backup has content
$backupFiles = @(Get-ChildItem $backupDir -Recurse -File -ErrorAction SilentlyContinue)
if ($backupFiles.Count -eq 0) {
    Show-Error "Backup directory is empty: $backupDir"
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
    Show-Warning "Edge is running ($($edgeProcesses.Count) processes)"
    Show-Info "Edge must be closed before restore"
    Write-Host ""

    if (-not (Confirm-Execution -Message "Terminate Edge and proceed with restore?")) {
        Write-Host ""
        Show-Info "Canceled"
        Write-Host ""
        return (New-ModuleResult -Status "Cancelled" -Message "User canceled (Edge running)")
    }

    Write-Host ""
    Show-Info "Terminating Edge processes..."

    try {
        Stop-Process -Name "msedge" -Force -ErrorAction Stop
        Start-Sleep -Seconds 2

        $remaining = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue)
        if ($remaining.Count -gt 0) {
            Show-Error "Edge processes still running after termination attempt"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Failed to terminate Edge")
        }

        Show-Success "Edge terminated"
    }
    catch {
        Show-Error "Failed to terminate Edge: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to terminate Edge: $_")
    }

    Write-Host ""
}
else {
    Show-Info "Edge is not running"
    Write-Host ""
}

# --- Confirm restore (with overwrite warning) ---
Show-Warning "This will OVERWRITE current Edge settings completely."
Write-Host ""
if (-not (Confirm-Execution -Message "Restore Edge profile from backup?")) {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# --- Execute Robocopy ---
Show-Info "Restore in progress (this may take a few minutes)..."
Write-Host ""

& robocopy.exe "$backupDir" "$targetDir" /MIR /XJ /MT /R:1 /W:1 /NFL /NDL
$exitCode = $LASTEXITCODE

Write-Host ""

# --- Result ---
if ($exitCode -lt 8) {
    Show-Separator
    Write-Host "Restore Results" -ForegroundColor Cyan
    Show-Separator
    Write-Host "  Status:   Success (robocopy exit: $exitCode)" -ForegroundColor Green
    Write-Host "  Restored: $sizeStr" -ForegroundColor White
    Show-Separator
    Write-Host ""

    return (New-ModuleResult -Status "Success" -Message "Restore completed ($sizeStr)")
}
else {
    Show-Error "Restore failed (robocopy exit code: $exitCode)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Robocopy failed (exit: $exitCode)")
}
