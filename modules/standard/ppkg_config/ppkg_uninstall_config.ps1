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
# - Restart-straddling flow: a locked/busy package routinely survives
#   the first run (entry and/or staged file). Such items report
#   Success with Verified=false ("pending") so AutoPilot/ErrorMode is
#   not tripped before __RESTART__, and the operational profile re-runs
#   this module after the restart. Locked files are registered for
#   delete-on-reboot while their path is still known. Fail is never
#   auto-escalated (the required number of reboots is package-dependent);
#   the checklist VERIFY column carries the pending signal instead.
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
$successCount  = 0
$skipCount     = 0
$failCount     = 0
$pendingCount  = 0
$verifiedCount = 0

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

    # Phase 1: Attempt cmdlet removal. A throw is NOT final here - on a
    # locked/busy package the first attempt routinely fails and the
    # operational flow re-runs this module across a restart. The final
    # verdict comes from the store re-query in Phase 3.
    $removeError = $null
    try {
        $null = Remove-ProvisioningPackage -PackageId $pkg.PackageId -ErrorAction Stop
        Show-Info "Remove-ProvisioningPackage completed for: $($item.PackageName)"
    }
    catch {
        $removeError = "$($_.Exception.Message)"
        Show-Warning "Remove-ProvisioningPackage failed: $_ (verdict by store re-query)"
    }

    # Phase 2: Delete the staged .ppkg file while its path is still
    # known (once the store entry is gone, later runs cannot find it).
    # If it stays locked, register a delete-on-reboot - a leftover
    # package file may embed credentials (Wi-Fi keys etc.) and must not
    # survive the kitting flow.
    $ppkgFilePath = $pkg.PackagePath
    $fileState = "none"   # none / deleted / scheduled / leftover
    if ($ppkgFilePath -and (Test-Path $ppkgFilePath)) {
        $maxRetry = 5
        for ($r = 0; $r -lt $maxRetry; $r++) {
            try {
                Remove-Item -Path $ppkgFilePath -Force -ErrorAction Stop
                Show-Info "Deleted package file: $ppkgFilePath"
                $fileState = "deleted"
                break
            }
            catch {
                if ($r -lt ($maxRetry - 1)) {
                    Show-Info "File locked, retrying in 2s... ($($r + 1)/$maxRetry)"
                    Start-Sleep -Seconds 2
                }
            }
        }

        if ($fileState -ne "deleted") {
            # MOVEFILE_DELAY_UNTIL_REBOOT (0x4): the OS deletes the file
            # early at next boot, before anything can lock it.
            if (-not ("PpkgFileCleanup" -as [type])) {
                Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class PpkgFileCleanup {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, uint dwFlags);
}
"@ -ErrorAction SilentlyContinue
            }
            $scheduled = $false
            try { $scheduled = [PpkgFileCleanup]::MoveFileEx($ppkgFilePath, $null, 0x4) } catch { }
            if ($scheduled) {
                Show-Warning "File locked - deletion scheduled at next restart: $ppkgFilePath"
                $fileState = "scheduled"
            }
            else {
                Show-Warning "Could not delete or schedule deletion (in use): $ppkgFilePath"
                $fileState = "leftover"
            }
        }
    }

    # Phase 3: Verdict by store re-query (after a short settle wait).
    # The previous flag-based result invented Success ("Cleaned up
    # package file") and Skip ("Already clean") for states where the
    # package was in fact still registered.
    Start-Sleep -Seconds 2
    $verifyErr = $null
    $verifyPkg = Get-ProvisioningPackage -AllInstalledPackages -ErrorAction SilentlyContinue -ErrorVariable verifyErr |
        Where-Object { $_.PackageName -eq $item.PackageName }

    # A failed re-query (not just "no match") must not be read as "package
    # gone". Treat a query error like a still-present package: pending, so
    # the operator re-runs after restart - the established straddle contract.
    if ($verifyPkg -or $verifyErr) {
        $why = if ($removeError)     { "removal attempt failed: $removeError" }
               elseif ($verifyErr)   { "verification re-query failed: $($verifyErr[0].Exception.Message)" }
               else                  { "entry still registered" }
        Show-Warning "Pending: $($item.PackageName) - $why (re-run this module after restart)"
        $successCount++
        $pendingCount++
    }
    else {
        switch ($fileState) {
            "scheduled" {
                Show-Success "Uninstalled: $($item.PackageName) (file deletion scheduled at next restart)"
                $pendingCount++
            }
            "leftover" {
                Show-Success "Uninstalled: $($item.PackageName) (WARNING: package file left on disk: $ppkgFilePath)"
                $pendingCount++
            }
            default {
                Show-Success "Uninstalled: $($item.PackageName) (PackageId: $($pkg.PackageId))"
                $verifiedCount++
            }
        }
        $successCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result Summary
# ========================================
# Verified: false while anything is pending (entry lingering or file
# cleanup delegated to reboot), true only when every processed item was
# confirmed clean, null when nothing was processed.
if ($pendingCount -gt 0) {
    return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
        -Title "PPKG Uninstall Results" -Verified $false `
        -MessageSuffix "($pendingCount pending - complete after restart)")
}
$verified = if ($verifiedCount -gt 0) { $true } else { $null }
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "PPKG Uninstall Results" -Verified $verified)
