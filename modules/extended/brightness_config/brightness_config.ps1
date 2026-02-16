# ========================================
# Brightness Configuration Script
# ========================================
# Sets display brightness via WMI
# (WmiMonitorBrightnessMethods).
# Note: Only supported on laptops / devices
# with built-in displays. Desktop PCs with
# external monitors are not supported.
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Brightness Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# 1. WMI Support Check
# ========================================
Write-Host "Checking WMI brightness support..." -ForegroundColor White

try {
    $monitor = Get-WmiObject -Namespace root\wmi -Class WmiMonitorBrightness -ErrorAction Stop
}
catch {
    $monitor = $null
}

if ($null -eq $monitor) {
    Write-Host "[SKIP] WMI brightness control not supported on this device (e.g., Desktop PC with external monitor)" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "WMI brightness not supported on this device")
}

Write-Host "[OK] WMI brightness control is available" -ForegroundColor Green
Write-Host ""

# ========================================
# 2. Get Current Brightness
# ========================================
$currentBrightness = $monitor.CurrentBrightness
Write-Host "[INFO] Current brightness: ${currentBrightness}%" -ForegroundColor Cyan
Write-Host ""

# ========================================
# 3. Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "brightness_list.csv"

$csvData = Import-CsvSafe -Path $csvPath -Description "brightness_list.csv"
if ($null -eq $csvData) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load brightness_list.csv")
}

if (-not (Test-CsvColumns -CsvData $csvData -RequiredColumns @("Enabled", "Brightness") -CsvName "brightness_list.csv")) {
    return (New-ModuleResult -Status "Error" -Message "brightness_list.csv missing required columns")
}

$enabledItems = @($csvData | Where-Object { $_.Enabled -eq "1" })

if ($enabledItems.Count -eq 0) {
    Write-Host "[INFO] No enabled entries in brightness_list.csv" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# 4. Validate & Display Targets
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Brightness Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasChanges = $false
$validItems = @()

foreach ($item in $enabledItems) {
    $brightness = [int]$item.Brightness
    $desc = if ($item.Description) { " $($item.Description)" } else { "" }

    # Validate range
    if ($brightness -lt 0 -or $brightness -gt 100) {
        Write-Host "  [ERROR] ${brightness}% is out of range (0-100)$desc" -ForegroundColor Red
        continue
    }

    $validItems += $item

    if ($currentBrightness -eq $brightness) {
        Write-Host "  [SKIP] ${brightness}%$desc (already set)" -ForegroundColor Gray
    }
    else {
        Write-Host "  [CHANGE] ${currentBrightness}% -> ${brightness}%$desc" -ForegroundColor White
        $hasChanges = $true
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($validItems.Count -eq 0) {
    Write-Host "[ERROR] No valid entries found" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No valid entries (all out of range)")
}

if (-not $hasChanges) {
    Write-Host "[INFO] Brightness already matches target value" -ForegroundColor Green
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Brightness already matches target")
}

# ========================================
# 5. Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Apply the above brightness setting?")) {
    Write-Host ""
    Write-Host "[INFO] Cancelled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User cancelled")
}

Write-Host ""

# ========================================
# 6. Apply Brightness
# ========================================
Write-Host "--- Applying Brightness ---" -ForegroundColor Cyan
Write-Host ""

$successCount = 0
$skipCount = 0
$errorCount = 0

foreach ($item in $validItems) {
    $brightness = [int]$item.Brightness
    $desc = if ($item.Description) { " ($($item.Description))" } else { "" }

    if ($currentBrightness -eq $brightness) {
        Write-Host "[SKIP] ${brightness}%$desc - already set" -ForegroundColor Gray
        $skipCount++
        continue
    }

    Write-Host "[INFO] Setting brightness to ${brightness}%$desc..." -ForegroundColor Cyan

    try {
        $methods = Get-WmiObject -Namespace root\wmi -Class WmiMonitorBrightnessMethods -ErrorAction Stop
        $methods.WmiSetBrightness(1, $brightness)

        Write-Host "[SUCCESS] Brightness changed to ${brightness}%" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to set brightness: $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }

    Write-Host ""
}

# ========================================
# 7. Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Brightness Configuration Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
if ($successCount -gt 0) {
    Write-Host "  Success: $successCount" -ForegroundColor Green
}
if ($skipCount -gt 0) {
    Write-Host "  Skipped: $skipCount (already set)" -ForegroundColor Gray
}
if ($errorCount -gt 0) {
    Write-Host "  Failed:  $errorCount" -ForegroundColor Red
}
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($errorCount -eq 0 -and $skipCount -gt 0 -and $successCount -eq 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    elseif ($errorCount -gt 0) { "Error" }
    else { "Success" }

return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $errorCount")
