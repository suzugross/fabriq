# ========================================
# Edge Profile Backup (with Auto Kill)
# ========================================
# Backs up Edge User Data directory using robocopy mirror.
# Automatically terminates Edge processes before backup.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Edge Profile Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# --- Paths ---
$sourceDir = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
$backupDir = Join-Path $PSScriptRoot "backup"

Write-Host "  Source:      $sourceDir" -ForegroundColor White
Write-Host "  Destination: $backupDir" -ForegroundColor White
Write-Host "  Mode:        MIRROR (Exact Copy)" -ForegroundColor Gray
Write-Host ""

# --- Source check ---
if (-not (Test-Path $sourceDir)) {
    Show-Error "Edge User Data not found: $sourceDir"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Edge User Data not found")
}

# --- Edge process check & kill ---
$edgeProcesses = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue)
if ($edgeProcesses.Count -gt 0) {
    Show-Warning "Edge is running ($($edgeProcesses.Count) processes)"
    Show-Info "Edge must be closed before backup"
    Write-Host ""

    $cancelResult = Confirm-ModuleExecution -Message "Terminate Edge and proceed with backup?"
    if ($null -ne $cancelResult) { return $cancelResult }

    Write-Host ""
    Show-Info "Terminating Edge processes..."

    try {
        Stop-Process -Name "msedge" -Force -ErrorAction Stop
        Start-Sleep -Seconds 2

        # Verify termination
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

# --- Confirm backup ---
$cancelResult = Confirm-ModuleExecution -Message "Start Edge profile backup?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# --- Create backup directory ---
if (-not (Test-Path $backupDir)) {
    try {
        $null = New-Item -ItemType Directory -Path $backupDir -Force
        Show-Info "Created backup directory"
    }
    catch {
        Show-Error "Failed to create backup directory: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create backup dir: $_")
    }
}

# --- Execute Robocopy ---
Show-Info "Backup in progress (this may take a few minutes)..."
Write-Host ""

& robocopy.exe "$sourceDir" "$backupDir" /MIR /XJ /MT /R:1 /W:1 /NFL /NDL
$exitCode = $LASTEXITCODE

Write-Host ""

# --- Result ---
if ($exitCode -lt 8) {
    # Calculate backup size
    $backupSize = 0
    try {
        $backupSize = (Get-ChildItem $backupDir -Recurse -File -ErrorAction SilentlyContinue |
                       Measure-Object -Property Length -Sum).Sum
    } catch {}
    $sizeStr = if ($backupSize -gt 1GB) {
        "{0:N1} GB" -f ($backupSize / 1GB)
    } elseif ($backupSize -gt 1MB) {
        "{0:N0} MB" -f ($backupSize / 1MB)
    } else {
        "{0:N0} KB" -f ($backupSize / 1KB)
    }

    Show-Separator
    Write-Host "Backup Results" -ForegroundColor Cyan
    Show-Separator
    Write-Host "  Status:    Success (robocopy exit: $exitCode)" -ForegroundColor Green
    Write-Host "  Size:      $sizeStr" -ForegroundColor White
    Write-Host "  Location:  $backupDir" -ForegroundColor White
    Show-Separator
    Write-Host ""

    return (New-ModuleResult -Status "Success" -Message "Backup completed ($sizeStr)")
}
else {
    Show-Error "Backup failed (robocopy exit code: $exitCode)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Robocopy failed (exit: $exitCode)")
}
