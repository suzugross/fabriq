# ========================================
# Start Layout Import Script
# ========================================
# Installs a provisioning package (.ppkg) that configures
# the Windows start menu layout. The PPKG must be built
# beforehand using the Start Layout Build module.
#
# [NOTES]
# - Requires administrator privileges
# - Uses Install-ProvisioningPackage -QuietInstall for silent operation
# - Input PPKG must exist under ppkg/ subdirectory
# ========================================

Write-Host ""
Show-Separator
Write-Host "Start Layout Import" -ForegroundColor Cyan
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
if (-not (Get-Command "Install-ProvisioningPackage" -ErrorAction SilentlyContinue)) {
    Show-Error "Install-ProvisioningPackage cmdlet is not available on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Install-ProvisioningPackage cmdlet not found")
}

$ppkgDir = Join-Path $PSScriptRoot "ppkg"

foreach ($item in $enabledItems) {
    $ppkgPath = Join-Path $ppkgDir "$($item.FileName).ppkg"
    if (-not (Test-Path $ppkgPath)) {
        Show-Error "PPKG file not found: $ppkgPath"
        Show-Error "Run Start Layout Build first to generate the PPKG file."
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "PPKG not found: $($item.FileName).ppkg")
    }
}

# ========================================
# Step 3: Pre-execution display
# ========================================
Show-Info "Import targets: $($enabledItems.Count) item(s)"
Write-Host ""

foreach ($item in $enabledItems) {
    $ppkgPath = Join-Path $ppkgDir "$($item.FileName).ppkg"

    # Check if already installed
    $installed = Get-ProvisioningPackage -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.FileName }

    if ($installed) {
        $marker = "[REINSTALL]"
        $markerColor = "Yellow"
    }
    else {
        $marker = "[INSTALL]"
        $markerColor = "White"
    }

    $ppkgSize = (Get-Item $ppkgPath).Length
    Write-Host "  [Id:$($item.Id)] $($item.FileName).ppkg  $marker" -ForegroundColor $markerColor
    Write-Host "    Path: $ppkgPath ($ppkgSize bytes)" -ForegroundColor DarkGray
    Write-Host ""
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Install provisioning package(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Import execution
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $ppkgPath = Join-Path $ppkgDir "$($item.FileName).ppkg"

    try {
        $null = Install-ProvisioningPackage -PackagePath $ppkgPath -QuietInstall -ForceInstall -ErrorAction Stop

        Show-Success "Installed: $($item.FileName).ppkg"
        $successCount++
    }
    catch {
        Show-Error "Failed to install: $($item.FileName).ppkg - $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Start Layout Import Results")
