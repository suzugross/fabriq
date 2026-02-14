# ========================================
# Store App Removal Script
# ========================================

Write-Host "Executing Store App removal process..." -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "storeapp_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] storeapp_list.csv not found: $csvPath" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "storeapp_list.csv not found")
}

try {
    $appList = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load storeapp_list.csv: $_" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load storeapp_list.csv: $_")
}

if ($appList.Count -eq 0) {
    Write-Host "[ERROR] storeapp_list.csv contains no data" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "storeapp_list.csv contains no data")
}

Write-Host "[INFO] Loaded $($appList.Count) app definitions" -ForegroundColor Cyan
Write-Host ""

# ========================================
# List Apps to Remove
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Store App Removal List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($app in $appList) {
    Write-Host "  $($app.AppName)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Delete the Store Apps listed above?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Removal Process
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($app in $appList) {
    $appName = $app.AppName
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Target: $appName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    $removed = $false

    # --- Remove AppxPackage (Current User) ---
    try {
        $appxPackage = Get-AppxPackage $appName -ErrorAction SilentlyContinue
        if ($appxPackage) {
            Remove-AppxPackage $appxPackage -ErrorAction Stop
            Write-Host "[SUCCESS] AppxPackage removed" -ForegroundColor Green
            $removed = $true
        }
        else {
            Write-Host "[INFO] AppxPackage not installed" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "[ERROR] Failed to remove AppxPackage: $_" -ForegroundColor Red
        $failCount++
        Write-Host ""
        continue
    }

    # --- Remove AppxProvisionedPackage (Provisioned) ---
    try {
        $provisionedPackage = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $appName }
        if ($provisionedPackage) {
            Remove-AppxProvisionedPackage -Online -PackageName $provisionedPackage.PackageName -ErrorAction Stop
            Write-Host "[SUCCESS] ProvisionedPackage removed" -ForegroundColor Green
            $removed = $true
        }
        else {
            Write-Host "[INFO] ProvisionedPackage not installed" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "[ERROR] Failed to remove ProvisionedPackage: $_" -ForegroundColor Red
        $failCount++
        Write-Host ""
        continue
    }

    if ($removed) {
        $successCount++
    }
    else {
        $skipCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Execution Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items" -ForegroundColor Yellow
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")