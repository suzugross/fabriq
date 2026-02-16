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
Show-Separator
Write-Host "Brightness Configuration" -ForegroundColor Cyan
Show-Separator
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
    Show-Skip "WMI brightness control not supported on this device (e.g., Desktop PC with external monitor)"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "WMI brightness not supported on this device")
}

Show-Success "WMI brightness control is available"
Write-Host ""

# ========================================
# 2. Get Current Brightness
# ========================================
$currentBrightness = $monitor.CurrentBrightness
Show-Info "Current brightness: ${currentBrightness}%"
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
    Show-Info "No enabled entries in brightness_list.csv"
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
        Show-Error "${brightness}% is out of range (0-100)$desc"
        continue
    }

    $validItems += $item

    if ($currentBrightness -eq $brightness) {
        Show-Skip "${brightness}%$desc (already set)"
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
    Show-Error "No valid entries found"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No valid entries (all out of range)")
}

if (-not $hasChanges) {
    Show-Skip "Brightness already matches target value"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Brightness already matches target")
}

# ========================================
# 5. Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Apply the above brightness setting?")) {
    Write-Host ""
    Show-Info "Cancelled"
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
        Show-Skip "${brightness}%$desc - already set"
        $skipCount++
        continue
    }

    Show-Info "Setting brightness to ${brightness}%$desc..."

    try {
        $methods = Get-WmiObject -Namespace root\wmi -Class WmiMonitorBrightnessMethods -ErrorAction Stop
        $methods.WmiSetBrightness(1, $brightness)

        Show-Success "Brightness changed to ${brightness}%"
        $successCount++
    }
    catch {
        Show-Error "Failed to set brightness: $($_.Exception.Message)"
        $errorCount++
    }

    Write-Host ""
}

# ========================================
# 7. Result Summary
# ========================================
Show-Separator
Write-Host "Brightness Configuration Results" -ForegroundColor Cyan
Show-Separator
if ($successCount -gt 0) {
    Write-Host "  Success: $successCount" -ForegroundColor Green
}
if ($skipCount -gt 0) {
    Write-Host "  Skipped: $skipCount (already set)" -ForegroundColor Gray
}
if ($errorCount -gt 0) {
    Write-Host "  Failed:  $errorCount" -ForegroundColor Red
}
Show-Separator
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($errorCount -eq 0 -and $skipCount -gt 0 -and $successCount -eq 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    elseif ($errorCount -gt 0) { "Error" }
    else { "Success" }

return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $errorCount")
