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
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Registry Import" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- CSV Loading ---
$csvPath = Join-Path $PSScriptRoot "reg_list.csv"
if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] reg_list.csv not found: $csvPath" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "reg_list.csv not found")
}

try {
    $allItems = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to read reg_list.csv: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to read reg_list.csv")
}

$items = @($allItems | Where-Object { $_.Enabled -eq "1" })
if ($items.Count -eq 0) {
    Write-Host "[INFO] No enabled entries in reg_list.csv" -ForegroundColor Gray
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# --- Backup directory check ---
$backupDir = Join-Path $PSScriptRoot "backup"
if (-not (Test-Path $backupDir)) {
    Write-Host "[ERROR] backup/ directory not found: $backupDir" -ForegroundColor Red
    Write-Host "[INFO] Run 'Registry Backup' first to create backup files." -ForegroundColor Yellow
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
Write-Host "[INFO] Import targets: $($importTargets.Count) registry keys" -ForegroundColor Cyan
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
    Write-Host "[ERROR] No backup files found for any entry" -ForegroundColor Red
    Write-Host "[INFO] Run 'Registry Backup' first to create backup files." -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No backup files found")
}

# --- Confirmation ---
Write-Host "[WARNING] This will merge registry data into the current system." -ForegroundColor Yellow
Write-Host ""
if (-not (Confirm-Execution -Message "Import the above registry backups?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

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
        Write-Host "  [SKIP] No backup file" -ForegroundColor Gray
        $skipCount++
        Write-Host ""
        continue
    }

    try {
        $process = Start-Process reg.exe -ArgumentList "import `"$($target.BackupFile.FullName)`"" -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -eq 0) {
            Write-Host "  [SUCCESS] Imported: $($target.BackupFile.Name)" -ForegroundColor Green
            $successCount++
        }
        else {
            Write-Host "  [ERROR] reg.exe exit code: $($process.ExitCode)" -ForegroundColor Red
            $failCount++
        }
    }
    catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $failCount++
    }

    Write-Host ""
}

# --- Summary ---
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Registry Import Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
if ($successCount -gt 0) {
    Write-Host "  Success: $successCount items" -ForegroundColor Green
}
if ($skipCount -gt 0) {
    Write-Host "  Skipped: $skipCount items (No backup)" -ForegroundColor Gray
}
if ($failCount -gt 0) {
    Write-Host "  Failed:  $failCount items" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Cyan

if ($successCount -gt 0) {
    Write-Host ""
    Write-Host "NOTE: Some changes may require sign-out or restart to take effect." -ForegroundColor Yellow
}
Write-Host ""

# --- ModuleResult ---
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($failCount -eq 0 -and $successCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($successCount -gt 0 -and $skipCount -gt 0) { "Success" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
