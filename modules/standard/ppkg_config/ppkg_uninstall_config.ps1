# ========================================
# PPKG Uninstall Script
# ========================================
# Uninstalls previously applied provisioning packages (.ppkg).
# Identifies target packages by matching PackageName from the CSV
# against installed packages via Get-ProvisioningPackage.
#
# [NOTES]
# - Requires administrator privileges
# - Idempotent: skips gracefully if the package is not installed
# ========================================

Write-Host ""
Show-Separator
Write-Host "PPKG Uninstall" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: CSV Loading
# ========================================
$csvPath = Join-Path $PSScriptRoot "ppkg_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "PackageName", "FileName", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load ppkg_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Prerequisites Check (Early Return)
# ========================================
if (-not (Get-Command "Get-ProvisioningPackage" -ErrorAction SilentlyContinue)) {
    Show-Error "Get-ProvisioningPackage cmdlet is not available on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Get-ProvisioningPackage cmdlet not found")
}

if (-not (Get-Command "Remove-ProvisioningPackage" -ErrorAction SilentlyContinue)) {
    Show-Error "Remove-ProvisioningPackage cmdlet is not available on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Remove-ProvisioningPackage cmdlet not found")
}


# ========================================
# Step 3: Pre-execution Display (Dry Run)
# ========================================
Show-Info "Uninstall targets: $($enabledItems.Count) item(s)"
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Provisioning Packages" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.PackageName }

    $pkg = Get-ProvisioningPackage -AllInstalledPackages -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.PackageName }

    if ($pkg) {
        $marker = "[INSTALLED]"
        $markerColor = "Yellow"
        Write-Host "  $marker $displayName" -ForegroundColor $markerColor
        Write-Host "    PackageName: $($item.PackageName)" -ForegroundColor DarkGray
        Write-Host "    PackageId:   $($pkg.PackageId)" -ForegroundColor DarkGray
    }
    else {
        $marker = "[NOT FOUND]"
        $markerColor = "DarkGray"
        Write-Host "  $marker $displayName" -ForegroundColor $markerColor
        Write-Host "    PackageName: $($item.PackageName)" -ForegroundColor DarkGray
    }

    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Uninstall provisioning package(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Uninstall Loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.PackageName }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Uninstalling: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Re-query at execution time (state may have changed since dry run)
    $pkg = Get-ProvisioningPackage -AllInstalledPackages -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.PackageName }

    if (-not $pkg) {
        Show-Skip "Not installed: $($item.PackageName)"
        $skipCount++
        Write-Host ""
        continue
    }

    # Phase 1: Attempt cmdlet removal
    $uninstallSuccess = $false
    try {
        $null = Remove-ProvisioningPackage -PackageId $pkg.PackageId -ErrorAction Stop
        Show-Info "Remove-ProvisioningPackage completed for: $($item.PackageName)"
        $uninstallSuccess = $true
    }
    catch {
        Show-Warning "Remove-ProvisioningPackage failed: $_ (proceeding to file cleanup)"
    }

    # Phase 2: Delete physical .ppkg file with retry
    $ppkgFilePath = $pkg.PackagePath
    $ppkgFileExists = ($ppkgFilePath -and (Test-Path $ppkgFilePath))
    $fileDeleted = -not $ppkgFileExists

    if ($ppkgFileExists) {
        $maxRetry = 5
        for ($r = 0; $r -lt $maxRetry; $r++) {
            try {
                Remove-Item -Path $ppkgFilePath -Force -ErrorAction Stop
                Show-Info "Deleted package file: $ppkgFilePath"
                $fileDeleted = $true
                break
            }
            catch {
                if ($r -lt ($maxRetry - 1)) {
                    Show-Info "File locked, retrying in 2s... ($($r + 1)/$maxRetry)"
                    Start-Sleep -Seconds 2
                }
            }
        }

        if (-not $fileDeleted) {
            Show-Warning "Could not delete package file (in use): $ppkgFilePath"
        }
    }

    # Phase 3: Determine result
    if ($uninstallSuccess) {
        Show-Success "Uninstalled: $($item.PackageName) (PackageId: $($pkg.PackageId))"
        $successCount++
    }
    elseif (-not $uninstallSuccess -and $fileDeleted -and $ppkgFileExists) {
        Show-Success "Cleaned up package file: $($item.PackageName)"
        $successCount++
    }
    elseif (-not $uninstallSuccess -and $fileDeleted -and -not $ppkgFileExists) {
        Show-Skip "Already clean: $($item.PackageName)"
        $skipCount++
    }
    else {
        Show-Error "Package removal failed: $($item.PackageName)"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "PPKG Uninstall Results")
