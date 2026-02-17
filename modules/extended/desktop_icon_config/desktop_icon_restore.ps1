# ========================================
# Desktop Icon Layout Restore
# ========================================
# Imports the latest desktop icon layout backup (.reg)
# from the backup/ directory to restore icon positions.
# Requires sign-out or restart to take effect.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Desktop Icon Layout Restore" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# --- Find backup files ---
$backupDir = Join-Path $PSScriptRoot "backup"

if (-not (Test-Path $backupDir)) {
    Show-Error "backup/ directory not found: $backupDir"
    Show-Info "Run 'Desktop Icon Backup' first."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "backup/ directory not found")
}

$backupFiles = @(Get-ChildItem -Path $backupDir -Filter "DesktopIcons_*.reg" -File -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending)

if ($backupFiles.Count -eq 0) {
    Show-Error "No backup files found in: $backupDir"
    Show-Info "Run 'Desktop Icon Backup' first."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No backup files found")
}

$latestBackup = $backupFiles[0]

# --- Display backup info ---
$fileSize = $latestBackup.Length
$sizeStr = if ($fileSize -gt 1KB) {
    "{0:N1} KB" -f ($fileSize / 1KB)
} else {
    "$fileSize bytes"
}

Write-Host "  Latest backup:" -ForegroundColor White
Write-Host "    File: $($latestBackup.Name)" -ForegroundColor White
Write-Host "    Date: $($latestBackup.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
Write-Host "    Size: $sizeStr" -ForegroundColor White

if ($backupFiles.Count -gt 1) {
    Write-Host "    ($($backupFiles.Count) backups available, using latest)" -ForegroundColor Gray
}

Write-Host ""

# --- Show all available backups if multiple ---
if ($backupFiles.Count -gt 1) {
    Write-Host "  Available backups:" -ForegroundColor Gray
    $showCount = [Math]::Min($backupFiles.Count, 5)
    for ($i = 0; $i -lt $showCount; $i++) {
        $marker = if ($i -eq 0) { " <- Latest" } else { "" }
        Write-Host "    [$($i + 1)] $($backupFiles[$i].Name) ($($backupFiles[$i].LastWriteTime.ToString('yyyy-MM-dd HH:mm')))$marker" -ForegroundColor DarkGray
    }
    if ($backupFiles.Count -gt 5) {
        Write-Host "    ... and $($backupFiles.Count - 5) more" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# --- Confirmation ---
Show-Warning "This will overwrite current desktop icon layout."
Show-Info "Changes take effect after sign-out or restart."
Write-Host ""
$cancelResult = Confirm-ModuleExecution -Message "Restore desktop icon layout from backup?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# --- Execute import ---
Show-Info "Importing registry data..."

try {
    $process = Start-Process reg.exe -ArgumentList "import `"$($latestBackup.FullName)`"" -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -eq 0) {
        Write-Host ""
        Show-Separator
        Write-Host "Restore Results" -ForegroundColor Cyan
        Show-Separator
        Write-Host "  Status:   Success" -ForegroundColor Green
        Write-Host "  File:     $($latestBackup.Name)" -ForegroundColor White
        Show-Separator
        Write-Host ""
        Show-Warning "Sign-out or restart required for changes to take effect."
        Write-Host ""

        return (New-ModuleResult -Status "Success" -Message "Restore completed (sign-out/restart required)")
    }
    else {
        Show-Error "reg.exe exit code: $($process.ExitCode)"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "reg.exe failed (exit: $($process.ExitCode))")
    }
}
catch {
    Show-Error "$($_.Exception.Message)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Import failed: $($_.Exception.Message)")
}
