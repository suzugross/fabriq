# ========================================
# Edge Profile Restore (with Auto Kill)
# ========================================
# Restores Edge folder to destinations defined in restore_dest.csv.
# Automatically terminates Edge processes before restore.
# WARNING: This overwrites current Edge settings completely.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Edge Profile Restore" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# --- Paths ---
$backupDir = Join-Path $PSScriptRoot "backup\Edge"
$csvPath   = Join-Path $PSScriptRoot "restore_dest.csv"

Write-Host "  Source:  $backupDir" -ForegroundColor White
Write-Host "  Dest CSV: $csvPath" -ForegroundColor White
Write-Host "  Mode:    MIRROR (Overwrite)" -ForegroundColor Gray
Write-Host ""

# --- Load destination CSV ---
$destList = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $destList -or $destList.Count -eq 0) {
    Show-Error "No enabled destinations found in: $csvPath"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No enabled destinations in restore_dest.csv")
}

Show-Info "Restore destinations ($($destList.Count)):"
foreach ($item in $destList) {
    $expanded = Expand-UserEnvironmentVariables $item.DestPath
    Write-Host "    -> $expanded\Edge  ($($item.Description))" -ForegroundColor Gray
}
Write-Host ""

# --- Backup existence check ---
if (-not (Test-Path $backupDir)) {
    Show-Error "Backup not found: $backupDir"
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

    $cancelResult = Confirm-ModuleExecution -Message "Terminate Edge and proceed with restore?"
    if ($null -ne $cancelResult) { return $cancelResult }

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
Show-Warning "This will OVERWRITE Edge settings at all destination paths."
Write-Host ""
$cancelResult = Confirm-ModuleExecution -Message "Restore Edge profile to all destinations?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# --- Execute Robocopy to each destination ---
$successCount = 0
$failCount    = 0

foreach ($item in $destList) {
    $expandedDest = Expand-UserEnvironmentVariables $item.DestPath
    $targetDir    = Join-Path $expandedDest "Edge"

    Show-Info "Restoring to: $targetDir"

    # Create destination directory if needed
    if (-not (Test-Path $targetDir)) {
        try {
            $null = New-Item -ItemType Directory -Path $targetDir -Force
        }
        catch {
            Show-Error "  Failed to create directory: $targetDir - $_"
            $failCount++
            continue
        }
    }

    & robocopy.exe "$backupDir" "$targetDir" /MIR /XJ /MT /R:1 /W:1 /NFL /NDL
    $exitCode = $LASTEXITCODE

    if ($exitCode -lt 8) {
        Show-Success "  OK (robocopy exit: $exitCode)"
        $successCount++
    }
    else {
        Show-Error "  Failed (robocopy exit: $exitCode)"
        $failCount++
    }

    Write-Host ""
}

# --- Result ---
Show-Separator
Write-Host "Restore Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Backup size: $sizeStr" -ForegroundColor White
Write-Host "  Success:     $successCount / $($destList.Count)" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Yellow" })
if ($failCount -gt 0) {
    Write-Host "  Failed:      $failCount" -ForegroundColor Red
}
Show-Separator
Write-Host ""

return (New-BatchResult -Success $successCount -Skip 0 -Fail $failCount -Title "Edge Restore")
