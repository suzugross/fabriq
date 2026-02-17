# ========================================
# Registry Deletion (HKCU + Default Profile)
# ========================================

$HIVE_PATH = "$env:SystemDrive\Users\Default\ntuser.dat"
$HIVE_KEY = "HKEY_USERS\Hive"

Write-Host ""
Show-Separator
Write-Host "Registry Deletion (HKCU + Default Profile)" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# Find CSV files
$csvFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter "reg_hkcu_list*.csv" -File | Sort-Object Name)

if ($csvFiles.Count -eq 0) {
    Show-Error "No files matching reg_hkcu_list*.csv found"
    return (New-ModuleResult -Status "Error" -Message "No files matching reg_hkcu_list*.csv found")
}

# Load CSV
$allItems = @()
$loadedFileCount = 0

foreach ($csvFile in $csvFiles) {
    $items = Import-CsvSafe -Path $csvFile.FullName -Description $csvFile.Name
    if ($null -ne $items) {
        $allItems += $items
        Show-Info "Loaded $($csvFile.Name) ($($items.Count) items)"
        $loadedFileCount++
    }
}

if ($loadedFileCount -eq 0) {
    Show-Error "Failed to load any CSV files"
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
    Show-Info "No valid registry settings found"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid registry settings found")
}

# ========================================
# List Deletions
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following registry values will be deleted" -ForegroundColor Yellow
Write-Host "  Target 1: HKEY_CURRENT_USER (Current User)" -ForegroundColor Yellow
Write-Host "  Target 2: Default Profile (For New Users)" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $regItems) {
    $checkPath = $item.'KeyPath' -replace '^HKEY_CURRENT_USER', 'HKCU:'
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
$cancelResult = Confirm-ModuleExecution -Message "Delete the above registry values?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""
Show-Info "Starting registry deletion..."
Write-Host ""

# ========================================
# Load Default Profile Hive
# ========================================
$hiveLoaded = $false

if (Test-Path $HIVE_PATH) {
    Show-Info "Loading Default Profile Hive..."
    & reg load $HIVE_KEY $HIVE_PATH 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Show-Success "Hive loaded: $HIVE_KEY"
        $hiveLoaded = $true
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
        }
    }
    else {
        Show-Error "Failed to load Hive"
        Show-Info "Deleting from HKCU only"
    }
}
else {
    Show-Error "ntuser.dat not found: $HIVE_PATH"
    Show-Info "Deleting from HKCU only"
}
Write-Host ""

# ========================================
# Execute Deletion
# ========================================
$successCount = 0
$failCount = 0

foreach ($item in $regItems) {
    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')" -ForegroundColor Yellow

    # ========================================
    # 1. Delete from HKCU
    # ========================================
    Write-Host "  [HKCU] Deleting from current user..." -ForegroundColor Gray

    $hkcuPath = $item.'KeyPath' -replace '^HKEY_CURRENT_USER', 'HKCU:'

    if (-not (Test-Path $hkcuPath)) {
        Write-Host "  [HKCU] Path not found (Skipped)" -ForegroundColor Gray
    }
    else {
        $existingValue = Get-ItemProperty -Path $hkcuPath -Name $item.'KeyName' -ErrorAction SilentlyContinue
        if (-not $existingValue) {
            Write-Host "  [HKCU] Value not found (Skipped)" -ForegroundColor Gray
        }
        else {
            try {
                Remove-ItemProperty -Path $hkcuPath -Name $item.'KeyName' -Force -ErrorAction Stop
                Write-Host "  [HKCU] Deleted" -ForegroundColor Green
            }
            catch {
                Write-Host "  [HKCU] Error: $_" -ForegroundColor Red
                $failCount++
                Write-Host ""
                continue
            }
        }
    }

    # ========================================
    # 2. Delete from Default Profile (HIVE)
    # ========================================
    if ($hiveLoaded) {
        Write-Host "  [HIVE] Deleting from Default Profile..." -ForegroundColor Gray

        $hivePsPath = $item.'KeyPath' -replace '^HKEY_CURRENT_USER', 'HKU:\Hive'

        if (-not (Test-Path $hivePsPath)) {
            Write-Host "  [HIVE] Path not found (Skipped)" -ForegroundColor Gray
        }
        else {
            $existingHiveValue = Get-ItemProperty -Path $hivePsPath -Name $item.'KeyName' -ErrorAction SilentlyContinue
            if (-not $existingHiveValue) {
                Write-Host "  [HIVE] Value not found (Skipped)" -ForegroundColor Gray
            }
            else {
                try {
                    Remove-ItemProperty -Path $hivePsPath -Name $item.'KeyName' -Force -ErrorAction Stop
                    Write-Host "  [HIVE] Deleted" -ForegroundColor Green
                }
                catch {
                    Write-Host "  [HIVE] Error: $_" -ForegroundColor Red
                    $failCount++
                    Write-Host ""
                    continue
                }
            }
        }
    }
    else {
        Write-Host "  [HIVE] Skipped (Hive not loaded)" -ForegroundColor Gray
    }

    $successCount++
    Write-Host ""
}

# ========================================
# Unload Hive
# ========================================
if ($hiveLoaded) {
    Show-Info "Unloading Hive..."

    if (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name HKU -Force
    }

    [gc]::Collect()
    Start-Sleep -Seconds 1

    & reg unload $HIVE_KEY 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Show-Success "Hive unloaded"
    }
    else {
        Show-Warning "Failed to unload Hive. Retrying..."
        Start-Sleep -Seconds 2
        [gc]::Collect()
        & reg unload $HIVE_KEY 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Show-Success "Hive unloaded (Retry success)"
        }
        else {
            Show-Error "Failed to unload Hive. Please unload manually."
        }
    }
    Write-Host ""
}

# Summary
return (New-BatchResult -Success $successCount -Fail $failCount -Title "Deletion Results")