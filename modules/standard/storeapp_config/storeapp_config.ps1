# ========================================
# Store App Removal Script
# ========================================

Show-Info "Executing Store App removal process..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "storeapp_list.csv"

$appList = Import-CsvSafe -Path $csvPath -Description "storeapp_list.csv"
if ($null -eq $appList -or $appList.Count -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load storeapp_list.csv")
}

Show-Info "Loaded $($appList.Count) app definitions"
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
    Show-Info "Canceled"
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
            Show-Success "AppxPackage removed"
            $removed = $true
        }
        else {
            Show-Info "AppxPackage not installed"
        }
    }
    catch {
        Show-Error "Failed to remove AppxPackage: $_"
        $failCount++
        Write-Host ""
        continue
    }

    # --- Remove AppxProvisionedPackage (Provisioned) ---
    try {
        $provisionedPackage = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $appName }
        if ($provisionedPackage) {
            Remove-AppxProvisionedPackage -Online -PackageName $provisionedPackage.PackageName -ErrorAction Stop
            Show-Success "ProvisionedPackage removed"
            $removed = $true
        }
        else {
            Show-Info "ProvisionedPackage not installed"
        }
    }
    catch {
        Show-Error "Failed to remove ProvisionedPackage: $_"
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
Show-Separator
Write-Host "Execution Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items" -ForegroundColor Yellow
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Show-Separator
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")