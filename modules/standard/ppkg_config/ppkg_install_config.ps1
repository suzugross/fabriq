# ========================================
# PPKG Install Script
# ========================================
# Installs provisioning packages (.ppkg) from the file/ directory
# using Install-ProvisioningPackage cmdlet. Targets are defined
# in ppkg_list.csv.
#
# [NOTES]
# - Requires administrator privileges
# - A reboot or re-logon may be required after installation
# ========================================

Write-Host ""
Show-Separator
Write-Host "PPKG Install" -ForegroundColor Cyan
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
if (-not (Get-Command "Install-ProvisioningPackage" -ErrorAction SilentlyContinue)) {
    Show-Error "Install-ProvisioningPackage cmdlet is not available on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Install-ProvisioningPackage cmdlet not found")
}

$fileDir = Join-Path $PSScriptRoot "file"
if (-not (Test-Path $fileDir)) {
    Show-Error "'file' directory not found: $fileDir"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "'file' directory not found")
}


# ========================================
# Step 3: Pre-execution Display (Dry Run)
# ========================================
Show-Info "Install targets: $($enabledItems.Count) item(s)"
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Provisioning Packages" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.PackageName }
    $ppkgPath = Join-Path $fileDir $item.FileName

    # File existence check
    if (-not (Test-Path $ppkgPath)) {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor Red
        Write-Host "    File: $ppkgPath" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    # Already installed check
    $installed = Get-ProvisioningPackage -ErrorAction SilentlyContinue |
        Where-Object { $_.PackageName -eq $item.PackageName }

    if ($installed) {
        $marker = "[REINSTALL]"
        $markerColor = "Yellow"
    }
    else {
        $marker = "[INSTALL]"
        $markerColor = "White"
    }

    $ppkgSize = (Get-Item $ppkgPath).Length
    Write-Host "  $marker $displayName" -ForegroundColor $markerColor
    Write-Host "    Package: $($item.PackageName)" -ForegroundColor DarkGray
    Write-Host "    File:    $($item.FileName) ($ppkgSize bytes)" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Install provisioning package(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Installation Loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.PackageName }
    $ppkgPath = Join-Path $fileDir $item.FileName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Installing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # File existence check
    if (-not (Test-Path $ppkgPath)) {
        Show-Error "File not found: $ppkgPath"
        $failCount++
        Write-Host ""
        continue
    }

    try {
        $null = Install-ProvisioningPackage -PackagePath $ppkgPath -QuietInstall -ForceInstall -ErrorAction Stop

        Show-Success "Installed: $($item.PackageName)"
        $successCount++
    }
    catch {
        Show-Error "Failed to install: $($item.PackageName) - $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "PPKG Install Results")
