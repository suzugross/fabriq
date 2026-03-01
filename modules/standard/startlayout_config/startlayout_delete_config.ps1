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

if (-not (Get-Command "Uninstall-ProvisioningPackage" -ErrorAction SilentlyContinue)) {
    Show-Error "Uninstall-ProvisioningPackage cmdlet is not available on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Uninstall-ProvisioningPackage cmdlet not found")
}

# ========================================
# Step 3: Pre-execution display
# ========================================
Show-Info "Delete targets: $($enabledItems.Count) item(s)"
Write-Host ""

foreach ($item in $enabledItems) {
    $pkg = Get-ProvisioningPackage -ErrorAction SilentlyContinue |
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
    $pkg = Get-ProvisioningPackage -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.FileName }

    if (-not $pkg) {
        Show-Skip "Not installed: $($item.FileName)"
        $skipCount++
        Write-Host ""
        continue
    }

    try {
        $null = Uninstall-ProvisioningPackage -PackageId $pkg.PackageId -ErrorAction Stop

        Show-Success "Uninstalled: $($item.FileName) (PackageId: $($pkg.PackageId))"
        $successCount++
    }
    catch {
        Show-Error "Failed to uninstall: $($item.FileName) - $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Start Layout Delete Results")
