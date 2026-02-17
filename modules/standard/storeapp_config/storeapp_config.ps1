# ========================================
# Store App Removal Script
# ========================================

Show-Info "Executing Store App removal process..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "storeapp_list.csv"

$appList = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $appList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load storeapp_list.csv")
}
if ($appList.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

$appList = @($appList | Sort-Object { [int]$_.No })
Write-Host ""

# ========================================
# List Apps to Remove
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Store App Removal List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($app in $appList) {
    Write-Host "  [$($app.No)] $($app.Description)" -ForegroundColor Yellow
    Write-Host "       $($app.AppName)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Delete the Store Apps listed above?"
if ($null -ne $cancelResult) { return $cancelResult }

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
    Write-Host "[$($app.No)] $($app.Description) ($appName)" -ForegroundColor Cyan
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
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")