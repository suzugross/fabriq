# ========================================
# Registry Deletion Script (HKLM)
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Registry Deletion (HKLM)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Find CSV files
$csvFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter "reg_hklm_list*.csv" -File | Sort-Object Name)

if ($csvFiles.Count -eq 0) {
    Write-Host "[ERROR] No files matching reg_hklm_list*.csv found" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "No files matching reg_hklm_list*.csv found")
}

# Load CSV
$allItems = @()
$loadedFileCount = 0

foreach ($csvFile in $csvFiles) {
    try {
        $items = @(Import-Csv -Path $csvFile.FullName -Encoding Default)
        $allItems += $items
        Write-Host "[INFO] Loaded $($csvFile.Name) ($($items.Count) items)" -ForegroundColor Cyan
        $loadedFileCount++
    }
    catch {
        Write-Host "[ERROR] Failed to load $($csvFile.Name): $_" -ForegroundColor Red
    }
}

if ($loadedFileCount -eq 0) {
    Write-Host "[ERROR] Failed to load any CSV files" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load any CSV files")
}

$regItems = @($allItems | Where-Object { $_.'Enabled' -eq '1' })
$skippedCount = $allItems.Count - $regItems.Count

Write-Host ""
Write-Host "[INFO] Total: $($regItems.Count) Enabled" -NoNewline -ForegroundColor Cyan
if ($skippedCount -gt 0) {
    Write-Host " / $skippedCount Skipped" -NoNewline -ForegroundColor Gray
}
Write-Host "" -ForegroundColor Cyan
Write-Host ""

if ($regItems.Count -eq 0) {
    Write-Host "[INFO] No valid registry settings found" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid registry settings found")
}

# ========================================
# List Deletions
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following registry values will be deleted" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $regItems) {
    $checkPath = $item.'KeyPath' -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
    $exists = Get-ItemProperty -Path $checkPath -Name $item.'KeyName' -ErrorAction SilentlyContinue

    $status = if ($exists) { "[Exists]" } else { "[Not Found]" }
    $statusColor = if ($exists) { "White" } else { "Gray" }

    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')  $status" -ForegroundColor $statusColor
    Write-Host "  Path: $($item.'KeyPath')"
    Write-Host "  Key:  $($item.'KeyName')"
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Confirmation
if (-not (Confirm-Execution -Message "Delete the above registry values?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""
Write-Host "[INFO] Starting registry deletion..." -ForegroundColor Cyan
Write-Host ""

# Execute Deletion
$successCount = 0
$skipCount = 0
$errorCount = 0

foreach ($item in $regItems) {
    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')" -ForegroundColor Yellow
    Write-Host "  Path: $($item.'KeyPath')"
    Write-Host "  Key:  $($item.'KeyName')"

    try {
        $regPath = $item.'KeyPath'
        $regPath = $regPath -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
        $regPath = $regPath -replace '^HKEY_CURRENT_USER', 'HKCU:'
        $regPath = $regPath -replace '^HKEY_CLASSES_ROOT', 'HKCR:'
        $regPath = $regPath -replace '^HKEY_USERS', 'HKU:'
        $regPath = $regPath -replace '^HKEY_CURRENT_CONFIG', 'HKCC:'

        # Check path existence
        if (-not (Test-Path $regPath)) {
            Write-Host "  [INFO] Path not found (Skipped)" -ForegroundColor Gray
            $skipCount++
            Write-Host ""
            continue
        }

        # Check value existence
        $existingValue = Get-ItemProperty -Path $regPath -Name $item.'KeyName' -ErrorAction SilentlyContinue
        if (-not $existingValue) {
            Write-Host "  [INFO] Value not found (Skipped)" -ForegroundColor Gray
            $skipCount++
            Write-Host ""
            continue
        }

        # Delete
        Remove-ItemProperty -Path $regPath -Name $item.'KeyName' -Force -ErrorAction Stop
        Write-Host "  [SUCCESS] Deleted" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "  [ERROR] $_" -ForegroundColor Red
        $errorCount++
    }

    Write-Host ""
}

# Summary
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Deletion Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Success: $successCount items" -ForegroundColor Green
if ($skipCount -gt 0) {
    Write-Host "Skipped: $skipCount items (Not found)" -ForegroundColor Gray
}
if ($errorCount -gt 0) {
    Write-Host "Failed: $errorCount items" -ForegroundColor Red
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    elseif ($errorCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $errorCount")