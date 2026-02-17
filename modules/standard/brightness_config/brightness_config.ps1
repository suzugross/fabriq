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
$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @("Enabled", "Brightness")
if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load brightness_list.csv")
}
if ($enabledItems.Count -eq 0) {
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
$cancelResult = Confirm-ModuleExecution -Message "Apply the above brightness setting?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# 6. Apply Brightness
# ========================================
Write-Host "--- Applying Brightness ---" -ForegroundColor Cyan
Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0

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
        $failCount++
    }

    Write-Host ""
}

# ========================================
# 7. Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Brightness Configuration Results")
