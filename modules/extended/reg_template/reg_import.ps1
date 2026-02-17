# ========================================
# Registry Import (Template)
# ========================================
# Imports .reg files from the backup/ directory.
# Uses the latest backup file for each registry path
# defined in reg_list.csv.
# Copy this directory and edit reg_list.csv to create
# custom registry import modules.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Registry Import" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# --- CSV Loading ---
$csvPath = Join-Path $PSScriptRoot "reg_list.csv"

$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) { return (New-ModuleResult -Status "Error" -Message "Failed to load reg_list.csv") }
if ($items.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }

# --- Backup directory check ---
$backupDir = Join-Path $PSScriptRoot "backup"
if (-not (Test-Path $backupDir)) {
    Show-Error "backup/ directory not found: $backupDir"
    Show-Info "Run 'Registry Backup' first to create backup files."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "backup/ directory not found")
}

# --- Find latest backup file for each entry ---
$importTargets = @()

foreach ($item in $items) {
    $regPath = $item.RegistryPath.Trim()
    $safeName = $regPath -replace '[\\/:*?"<>|]', '_'
    if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0, 80) }

    # Find matching .reg files (pattern: *_safeName.reg)
    $matchingFiles = @(Get-ChildItem -Path $backupDir -Filter "*_${safeName}.reg" -File -ErrorAction SilentlyContinue |
                       Sort-Object LastWriteTime -Descending)

    $importTargets += [PSCustomObject]@{
        RegistryPath = $regPath
        Description  = $item.Description
        BackupFile   = if ($matchingFiles.Count -gt 0) { $matchingFiles[0] } else { $null }
        BackupCount  = $matchingFiles.Count
    }
}

# --- Display target list ---
Show-Info "Import targets: $($importTargets.Count) registry keys"
Write-Host ""

$index = 0
foreach ($target in $importTargets) {
    $index++

    if ($null -ne $target.BackupFile) {
        $marker = "[Ready]"
        $markerColor = "Green"
        Write-Host "  [$index] $($target.RegistryPath)  $marker" -ForegroundColor $markerColor
        Write-Host "      File: $($target.BackupFile.Name) ($($target.BackupFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))" -ForegroundColor White
        if ($target.BackupCount -gt 1) {
            Write-Host "      ($($target.BackupCount) backups available, using latest)" -ForegroundColor Gray
        }
    }
    else {
        $marker = "[Missing]"
        $markerColor = "Red"
        Write-Host "  [$index] $($target.RegistryPath)  $marker" -ForegroundColor $markerColor
        Write-Host "      No backup file found" -ForegroundColor Red
    }

    if ($target.Description) {
        Write-Host "      $($target.Description)" -ForegroundColor Gray
    }
    Write-Host ""
}

# Check if any files are available
$readyTargets = @($importTargets | Where-Object { $null -ne $_.BackupFile })
if ($readyTargets.Count -eq 0) {
    Show-Error "No backup files found for any entry"
    Show-Info "Run 'Registry Backup' first to create backup files."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No backup files found")
}

# --- Confirmation ---
Show-Warning "This will merge registry data into the current system."
Write-Host ""
$cancelResult = Confirm-ModuleExecution -Message "Import the above registry backups?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# --- Execute import ---
$successCount = 0
$skipCount = 0
$failCount = 0
$total = $importTargets.Count
$current = 0

foreach ($target in $importTargets) {
    $current++

    Write-Host "[$current/$total] $($target.RegistryPath)" -ForegroundColor Cyan

    if ($null -eq $target.BackupFile) {
        Show-Skip "No backup file"
        $skipCount++
        Write-Host ""
        continue
    }

    try {
        $process = Start-Process reg.exe -ArgumentList "import `"$($target.BackupFile.FullName)`"" -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -eq 0) {
            Show-Success "Imported: $($target.BackupFile.Name)"
            $successCount++
        }
        else {
            Show-Error "reg.exe exit code: $($process.ExitCode)"
            $failCount++
        }
    }
    catch {
        Show-Error "$($_.Exception.Message)"
        $failCount++
    }

    Write-Host ""
}

# --- Summary ---
if ($successCount -gt 0) {
    Show-Warning "Some changes may require sign-out or restart to take effect."
    Write-Host ""
}
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Registry Import Results")
