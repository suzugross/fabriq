# ========================================
# Start Layout Delete Script
# ========================================
# Uninstalls a previously applied start layout provisioning
# package. Identifies the target package by matching the
# PackageName against the FileName from the CSV.
#
# [NOTES]
# - Requires administrator privileges
# - Idempotent: skips gracefully if the package is not installed
# ========================================

Write-Host ""
Show-Separator
Write-Host "Start Layout Delete" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "startlayout_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Id", "FileName")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load startlayout_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Pre-flight checks
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
# Step 3: Pre-execution display
# ========================================
Show-Info "Delete targets: $($enabledItems.Count) item(s)"
Write-Host ""

foreach ($item in $enabledItems) {
    $pkg = Get-ProvisioningPackage -AllInstalledPackages -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.FileName }

    if ($pkg) {
        $marker = "[INSTALLED]"
        $markerColor = "Yellow"
        Write-Host "  [Id:$($item.Id)] $($item.FileName)  $marker" -ForegroundColor $markerColor
        Write-Host "    PackageId: $($pkg.PackageId)" -ForegroundColor DarkGray
    }
    else {
        $marker = "[NOT FOUND]"
        $markerColor = "DarkGray"
        Write-Host "  [Id:$($item.Id)] $($item.FileName)  $marker" -ForegroundColor $markerColor
    }

    Write-Host ""
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Uninstall provisioning package(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Delete execution
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    # Re-query to get current state at execution time
    $pkg = Get-ProvisioningPackage -AllInstalledPackages -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.FileName }

    if (-not $pkg) {
        Show-Skip "Not installed: $($item.FileName)"
        $skipCount++
        Write-Host ""
        continue
    }

    try {
        # Phase 1: Attempt cmdlet removal
        $null = Remove-ProvisioningPackage -PackageId $pkg.PackageId -ErrorAction Stop
        Show-Info "Remove-ProvisioningPackage completed for: $($item.FileName)"
    }
    catch {
        Show-Warning "Remove-ProvisioningPackage failed: $_ (proceeding to file cleanup)"
    }

    # Phase 2: Delete physical .ppkg file if it still exists
    $ppkgFilePath = $pkg.PackagePath
    if ($ppkgFilePath -and (Test-Path $ppkgFilePath)) {
        try {
            Remove-Item -Path $ppkgFilePath -Force -ErrorAction Stop
            Show-Info "Deleted package file: $ppkgFilePath"
        }
        catch {
            Show-Error "Failed to delete package file: $ppkgFilePath - $_"
            $failCount++
            Write-Host ""
            continue
        }
    }

    # Phase 3: Verify removal
    $verifyPkg = Get-ProvisioningPackage -AllInstalledPackages -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.FileName }

    if ($verifyPkg) {
        Show-Error "Package still exists after removal: $($item.FileName)"
        $failCount++
    }
    else {
        Show-Success "Uninstalled: $($item.FileName) (PackageId: $($pkg.PackageId))"
        $successCount++
    }

    Write-Host ""
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Start Layout Delete Results")
