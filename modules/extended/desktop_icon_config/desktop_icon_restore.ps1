# ========================================
# Desktop Icon Layout Restore
# ========================================
# Imports the latest desktop icon layout backup (.reg)
# from the backup/ directory to restore icon positions.
# Requires sign-out or restart to take effect.
# ========================================

# Resolve logged-on user's HKCU target
$hkcuInfo = Resolve-HkcuRoot

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
if ($hkcuInfo.Redirected) {
    Write-Host "  Target:       $($hkcuInfo.Label)" -ForegroundColor Magenta
}

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

$importFile = $latestBackup.FullName
$tempFileCreated = $false

# When redirected, rewrite .reg to target logged-on user's hive via HKU
if ($hkcuInfo.Redirected) {
    Show-Info "Redirecting import to logged-on user ($($hkcuInfo.Label))..."
    $regContent = Get-Content -Path $latestBackup.FullName -Raw -Encoding Unicode
    $regContent = $regContent -replace 'HKEY_CURRENT_USER', "HKEY_USERS\$($hkcuInfo.SID)"
    $importFile = Join-Path $env:TEMP "fabriq_desktop_restore_temp.reg"
    Set-Content -Path $importFile -Value $regContent -Encoding Unicode -NoNewline
    $tempFileCreated = $true
}

try {
    $process = Start-Process reg.exe -ArgumentList "import `"$importFile`"" -Wait -PassThru -NoNewWindow
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
finally {
    if ($tempFileCreated -and (Test-Path $importFile)) {
        Remove-Item -Path $importFile -Force -ErrorAction SilentlyContinue
    }
}
