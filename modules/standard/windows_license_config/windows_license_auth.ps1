# ========================================
# Windows License Activation Script
# ========================================
# Description: Activates the Windows license
# by triggering RefreshLicenseStatus.
# Skips if already activated (idempotent).
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "  Activate Windows License" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper: License Status Text
# ========================================
function Get-LicenseStatusText {
    param([int]$Status)
    switch ($Status) {
        0 { "Unlicensed" }
        1 { "Licensed" }
        2 { "OOBE Grace Period" }
        3 { "Out of Tolerance" }
        4 { "Non-Genuine Grace Period" }
        5 { "Notification Mode" }
        6 { "Extended Grace Period" }
        default { "Unknown ($Status)" }
    }
}

# ========================================
# Step 1: Check Current License Status
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Current License Status" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

$osLicense = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
             Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } |
             Select-Object -First 1

if ($null -eq $osLicense) {
    Show-Error "No Windows product key found on this system."
    Show-Info "Please install a product key first."
    Write-Host "----------------------------------------" -ForegroundColor White
    return (New-ModuleResult -Status "Error" -Message "No product key installed")
}

$editionName = $osLicense.Name.Split(',')[0]
Write-Host "  Edition:        $editionName"
Write-Host "  Partial Key:    $($osLicense.PartialProductKey)"
Write-Host "  License Status: $(Get-LicenseStatusText $osLicense.LicenseStatus)"
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Step 2: Idempotency Check
# ========================================
if ($osLicense.LicenseStatus -eq 1) {
    Show-Skip "Windows is already activated."
    return (New-ModuleResult -Status "Skipped" -Message "Already activated ($editionName)")
}

# ========================================
# Step 3: Confirm
# ========================================
if (-not (Confirm-Execution -Message "Activate Windows license?")) {
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

# ========================================
# Step 4: Trigger Activation
# ========================================
Write-Host ""
Show-Info "Triggering Windows activation..."

try {
    $service = Get-CimInstance -ClassName SoftwareLicensingService
    $null = Invoke-CimMethod -InputObject $service -MethodName RefreshLicenseStatus
    Write-Host "[INFO] Activation request sent. Waiting for result..." -ForegroundColor Gray
}
catch {
    Show-Error "Failed to trigger activation: $($_.Exception.Message)"
    return (New-ModuleResult -Status "Error" -Message "Activation failed: $($_.Exception.Message)")
}

# ========================================
# Step 5: Verify Result
# ========================================
Start-Sleep -Seconds 3

$finalLicense = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
                Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } |
                Select-Object -First 1

Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host "Activation Result" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

if ($null -eq $finalLicense) {
    Show-Error "Could not retrieve license status"
    Write-Host "========================================" -ForegroundColor White
    return (New-ModuleResult -Status "Error" -Message "Could not verify activation")
}

$statusText = Get-LicenseStatusText $finalLicense.LicenseStatus
$statusColor = switch ($finalLicense.LicenseStatus) {
    1 { "Green" }
    0 { "Red" }
    default { "Yellow" }
}

$finalEdition = $finalLicense.Name.Split(',')[0]
Write-Host "  Edition:        $finalEdition"
Write-Host "  Partial Key:    $($finalLicense.PartialProductKey)"
Write-Host "  License Status: $statusText" -ForegroundColor $statusColor
Write-Host "========================================" -ForegroundColor White

if ($finalLicense.LicenseStatus -eq 1) {
    return (New-ModuleResult -Status "Success" -Message "Activated ($finalEdition)")
}
else {
    return (New-ModuleResult -Status "Error" -Message "Activation incomplete ($statusText)")
}
