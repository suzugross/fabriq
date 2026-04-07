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
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying removed apps..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($app in $appList) {
    $appName = $app.AppName
    $displayName = $app.Description

    $appxRemains = $null -ne (Get-AppxPackage $appName -ErrorAction SilentlyContinue)
    $provRemains = $null -ne (Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -eq $appName })

    if (-not $appxRemains -and -not $provRemains) {
        Write-Host "  [VERIFIED] $displayName (not installed)" -ForegroundColor Green
        $verifyPass++
    } else {
        $remaining = @()
        if ($appxRemains) { $remaining += "AppxPackage" }
        if ($provRemains) { $remaining += "Provisioned" }
        Write-Host "  [VERIFY FAILED] $displayName (still: $($remaining -join ', '))" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results" -Verified $verified)