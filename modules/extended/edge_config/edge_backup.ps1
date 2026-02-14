# ========================================
# Edge Profile Backup (with Auto Kill)
# ========================================
# Backs up Edge User Data directory using robocopy mirror.
# Automatically terminates Edge processes before backup.
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Edge Profile Backup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
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
    Write-Host "[ERROR] Edge User Data not found: $sourceDir" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Edge User Data not found")
}

# --- Edge process check & kill ---
$edgeProcesses = @(Get-Process -Name "msedge" -ErrorAction SilentlyContinue)
if ($edgeProcesses.Count -gt 0) {
    Write-Host "[WARNING] Edge is running ($($edgeProcesses.Count) processes)" -ForegroundColor Yellow
    Write-Host "[INFO] Edge must be closed before backup" -ForegroundColor Yellow
    Write-Host ""

    if (-not (Confirm-Execution -Message "Terminate Edge and proceed with backup?")) {
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

        # Verify termination
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

# --- Confirm backup ---
if (-not (Confirm-Execution -Message "Start Edge profile backup?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# --- Create backup directory ---
if (-not (Test-Path $backupDir)) {
    try {
        $null = New-Item -ItemType Directory -Path $backupDir -Force
        Write-Host "[INFO] Created backup directory" -ForegroundColor Gray
    }
    catch {
        Write-Host "[ERROR] Failed to create backup directory: $_" -ForegroundColor Red
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create backup dir: $_")
    }
}

# --- Execute Robocopy ---
Write-Host "[INFO] Backup in progress (this may take a few minutes)..." -ForegroundColor Cyan
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

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Backup Results" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Status:    Success (robocopy exit: $exitCode)" -ForegroundColor Green
    Write-Host "  Size:      $sizeStr" -ForegroundColor White
    Write-Host "  Location:  $backupDir" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    return (New-ModuleResult -Status "Success" -Message "Backup completed ($sizeStr)")
}
else {
    Write-Host "[ERROR] Backup failed (robocopy exit code: $exitCode)" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Robocopy failed (exit: $exitCode)")
}
