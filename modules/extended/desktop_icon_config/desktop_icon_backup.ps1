# ========================================
# Desktop Icon Layout Backup
# ========================================
# Exports the desktop icon layout registry key
# (HKCU\Software\Microsoft\Windows\Shell\Bags\1\Desktop)
# to a .reg file for later restoration.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Desktop Icon Layout Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""

$registryPath = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Bags\1\Desktop'

# --- Check if registry key exists ---
$checkPath = 'HKCU:\Software\Microsoft\Windows\Shell\Bags\1\Desktop'
$keyExists = Test-Path $checkPath -ErrorAction SilentlyContinue

Write-Host "  Registry Key: $registryPath" -ForegroundColor White
if ($keyExists) {
    Write-Host "  Status:       Exists [OK]" -ForegroundColor Green
}
else {
    Write-Host "  Status:       Not Found [--]" -ForegroundColor Yellow
}
Write-Host ""

Show-Info "Arrange your desktop icons in the desired layout"
Write-Host "       before proceeding with the backup." -ForegroundColor Yellow
Write-Host ""

# --- Confirmation ---
if (-not (Confirm-Execution -Message "Backup desktop icon layout?")) {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# --- Backup directory ---
$backupDir = Join-Path $PSScriptRoot "backup"
if (-not (Test-Path $backupDir)) {
    try {
        $null = New-Item -ItemType Directory -Path $backupDir -Force
    }
    catch {
        Show-Error "Failed to create backup directory: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create backup dir")
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportFile = Join-Path $backupDir "DesktopIcons_${timestamp}.reg"

# --- Execute backup ---
Show-Info "Exporting registry key..."

try {
    $process = Start-Process reg.exe -ArgumentList "export `"$registryPath`" `"$exportFile`" /y" -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -eq 0) {
        # Get file size
        $fileSize = (Get-Item $exportFile -ErrorAction SilentlyContinue).Length
        $sizeStr = if ($fileSize -gt 1KB) {
            "{0:N1} KB" -f ($fileSize / 1KB)
        } else {
            "$fileSize bytes"
        }

        Write-Host ""
        Show-Separator
        Write-Host "Backup Results" -ForegroundColor Cyan
        Show-Separator
        Write-Host "  Status:   Success" -ForegroundColor Green
        Write-Host "  File:     $([System.IO.Path]::GetFileName($exportFile))" -ForegroundColor White
        Write-Host "  Size:     $sizeStr" -ForegroundColor White
        Write-Host "  Location: $backupDir" -ForegroundColor White
        Show-Separator
        Write-Host ""

        return (New-ModuleResult -Status "Success" -Message "Backup completed ($sizeStr)")
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
    return (New-ModuleResult -Status "Error" -Message "Export failed: $($_.Exception.Message)")
}
