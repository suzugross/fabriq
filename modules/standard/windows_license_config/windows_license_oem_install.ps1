# ========================================
# Windows OEM Product Key Installation Script
# ========================================
# [PURPOSE]
# Reads the OEM product key stored in this PC's firmware
# (SoftwareLicensingService.OA3xOriginalProductKey) and installs it as the
# Windows product key. There is no CSV: the key always comes from the PC
# the script runs on.
#
# [NOTES]
# - Requires administrator privileges.
# - A PC without a firmware key (virtual machines, volume-license and
#   self-built PCs) ends in Error. There is deliberately no fallback to
#   license_key.csv or to manual input, so Success and Skipped always
#   refer to the OEM key of this PC and never to a substitute key.
# - Skips when the OEM key is already the installed key (idempotent).
# - Installs the key only. Activation is windows_license_auth.ps1.
# - The key is never displayed or logged in full (Get-MaskedKey). The raw
#   key is handed to the helper functions and to InstallProductKey only.
# - Get-MaskedKey / Get-LicenseStatusText and the install and read-back
#   steps mirror windows_license_install.ps1. Keep the two in sync.
# ========================================


# ========================================
# Helper functions
# ========================================
# Every function below is pure (no access to the PC) and is unit-tested in
# tests/modules/windows_license_config.
# ========================================

# Returns a masked form keeping only the last 5 characters visible.
# Dashes are preserved so the standard 5-5-5-5-5 layout remains
# recognizable. Falls back to length-only masking for non-conforming
# inputs. Used to avoid leaking raw keys to the PowerShell transcript
# via Write-Host / Show-* paths.
function Get-MaskedKey {
    param([string]$Key)
    if ([string]::IsNullOrEmpty($Key)) { return '' }
    $len = $Key.Length
    if ($len -le 5) { return ('*' * $len) }
    $sb = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $len - 5; $i++) {
        if ($Key[$i] -eq '-') { [void]$sb.Append('-') } else { [void]$sb.Append('*') }
    }
    [void]$sb.Append($Key.Substring($len - 5))
    return $sb.ToString()
}

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

# Last 5 characters of a product key, upper-cased: the form Windows reports
# as PartialProductKey. Empty when the key is shorter than that.
function Get-KeyPartial {
    param([string]$Key)
    $trimmed = "$Key".Trim()
    if ($trimmed.Length -lt 5) { return '' }
    return $trimmed.Substring($trimmed.Length - 5).ToUpperInvariant()
}

# Drops blank entries and normalizes the rest for comparison. Callers wrap
# the result in @() (an empty result unrolls to $null on return).
function Get-NormalizedPartials {
    param([string[]]$Partials)
    return @($Partials |
             Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
             ForEach-Object { $_.Trim().ToUpperInvariant() })
}

# Decides what to do with the firmware key:
#   NoKey            - the firmware holds no key
#   BadFormat        - the value is not XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
#   AlreadyInstalled - an installed Windows key ends in the same 5 characters
#   Install          - everything else. This includes "the installed keys
#                      could not be read": an unreadable state is never
#                      taken as "already installed".
function Get-OemKeyPlan {
    param(
        [string]$OemKey,
        [string[]]$InstalledPartials
    )
    $key = "$OemKey".Trim()
    if ([string]::IsNullOrWhiteSpace($key)) { return 'NoKey' }
    if ($key -notmatch '^[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}$') {
        return 'BadFormat'
    }
    $installed = @(Get-NormalizedPartials -Partials $InstalledPartials)
    if ($installed -contains (Get-KeyPartial $key)) { return 'AlreadyInstalled' }
    return 'Install'
}

# Classifies the read-back after InstallProductKey:
#   Match      - an installed Windows key ends in the expected 5 characters
#   Mismatch   - keys are readable but none is the expected one (the old
#                key is still in effect)
#   Unreadable - no installed key could be read at all
function Get-OemVerifyOutcome {
    param(
        [string]$ExpectedPartial,
        [string[]]$ActualPartials
    )
    $actual = @(Get-NormalizedPartials -Partials $ActualPartials)
    if ($actual.Count -eq 0) { return 'Unreadable' }
    if ($actual -contains "$ExpectedPartial".Trim().ToUpperInvariant()) { return 'Match' }
    return 'Mismatch'
}


# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "  Install OEM Product Key" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Read the OEM Key from Firmware
# ========================================
$service = $null
try {
    $service = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop
}
catch {
    Show-Error "Failed to query the licensing service: $($_.Exception.Message)"
    return (New-ModuleResult -Status "Error" -Message "Licensing service query failed: $($_.Exception.Message)")
}

if ($null -eq $service) {
    Show-Error "The licensing service returned no data."
    return (New-ModuleResult -Status "Error" -Message "Licensing service returned no data")
}

# [string] cast: the property is $null on a PC without a firmware key, which
# must reach the NoKey branch below, not crash on .Trim()
$oemKey  = ([string]$service.OA3xOriginalProductKey).Trim()
$oemDesc = ([string]$service.OA3xOriginalProductKeyDescription).Trim()

# ========================================
# Step 2: Read Current License Status and Decide
# ========================================
$winProducts = @(Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
                 Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" })
$installedPartials = @($winProducts | ForEach-Object { "$($_.PartialProductKey)" })

$plan = Get-OemKeyPlan -OemKey $oemKey -InstalledPartials $installedPartials

if ($plan -eq 'NoKey') {
    Show-Error "No OEM product key found in the firmware of this PC."
    Show-Info "Virtual machines and PCs sold without a preinstalled Windows license have no firmware key."
    return (New-ModuleResult -Status "Error" -Message "No OEM product key in firmware")
}

if ($plan -eq 'BadFormat') {
    Show-Error "The firmware OEM key has an unexpected format (expected XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
    return (New-ModuleResult -Status "Error" -Message "Firmware OEM key has an unexpected format")
}

$expectedPartial = Get-KeyPartial $oemKey

# ========================================
# Step 3: Show the Firmware Key and the Current License Status
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Firmware OEM Key" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "  Key:            $(Get-MaskedKey $oemKey)" -ForegroundColor Yellow
if ($oemDesc) {
    Write-Host "  Description:    $oemDesc" -ForegroundColor Gray
}
Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Current License Status" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

if ($winProducts.Count -gt 0) {
    foreach ($product in $winProducts) {
        $editionName = $product.Name.Split(',')[0]
        Write-Host "  Edition:        $editionName"
        Write-Host "  Partial Key:    $($product.PartialProductKey)"
        Write-Host "  License Status: $(Get-LicenseStatusText $product.LicenseStatus)"
    }
}
else {
    Write-Host "  No existing product key found" -ForegroundColor Gray
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Step 4: Idempotency Check
# ========================================
if ($plan -eq 'AlreadyInstalled') {
    Show-Skip "The OEM key is already installed."
    return (New-ModuleResult -Status "Skipped" -Message "OEM key already installed (Partial: $expectedPartial)" -Verified $true)
}

# ========================================
# Step 5: Confirm
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Install the OEM product key from firmware?"
if ($null -ne $cancelResult) { return $cancelResult }

# ========================================
# Step 6: Install Product Key
# ========================================
Write-Host ""
try {
    # InstallProductKey overwrites any existing key; Step 7 verifies the
    # read-back.
    Show-Info "Installing the OEM product key..."
    $null = Invoke-CimMethod -InputObject $service -MethodName InstallProductKey `
        -Arguments @{ProductKey = $oemKey} -ErrorAction Stop
    Show-Success "Product key installed"
}
catch {
    Show-Error "Failed to install the OEM product key: $($_.Exception.Message)"
    if ($oemDesc) {
        # The usual cause is an edition mismatch (e.g. a Home key in the
        # firmware of a PC imaged with Pro). Put both facts side by side.
        Show-Info "The firmware key is for: $oemDesc - check that it matches the installed Windows edition."
    }
    return (New-ModuleResult -Status "Error" -Message "Install failed: $($_.Exception.Message)")
}

# ========================================
# Step 7: Verify Installation (read-back against the firmware key)
# ========================================
# Mere existence of a Windows product with a PartialProductKey is not
# verification - a leftover OLD key looks identical. Match the last 5
# characters of the firmware key (the part Get-MaskedKey keeps visible
# by design, so no new exposure). Retry to absorb WMI refresh latency.
$outcome        = 'Unreadable'
$actualPartials = @()
$winProducts    = @()

for ($attempt = 1; $attempt -le 3; $attempt++) {
    Start-Sleep -Seconds 2
    $winProducts = @(Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
                     Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" })
    $actualPartials = @($winProducts | ForEach-Object { "$($_.PartialProductKey)" })
    $outcome = Get-OemVerifyOutcome -ExpectedPartial $expectedPartial -ActualPartials $actualPartials
    if ($outcome -eq 'Match') { break }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host "Installation Result" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

if ($outcome -eq 'Match') {
    $matchedProduct = $winProducts |
        Where-Object { "$($_.PartialProductKey)".ToUpper() -eq $expectedPartial } |
        Select-Object -First 1
    $editionName = $matchedProduct.Name.Split(',')[0]
    Write-Host "  Edition:        $editionName" -ForegroundColor Green
    Write-Host "  Partial Key:    $($matchedProduct.PartialProductKey) (matches firmware key)" -ForegroundColor Green
    Write-Host "  License Status: $(Get-LicenseStatusText $matchedProduct.LicenseStatus)"
    Write-Host "========================================" -ForegroundColor White

    return (New-ModuleResult -Status "Success" -Message "OEM key installed (Partial: $($matchedProduct.PartialProductKey))" -Verified $true)
}

if ($outcome -eq 'Mismatch') {
    # Products are readable but none carries the firmware key - the OLD
    # key is still in effect. Definite contradiction -> fail closed.
    $actualText = $actualPartials -join ', '
    Show-Error "Installed key mismatch (expected last5: $expectedPartial, actual: $actualText)"
    Write-Host "========================================" -ForegroundColor White

    return (New-ModuleResult -Status "Error" -Message "Key mismatch after install (expected: $expectedPartial, actual: $actualText)")
}

# No licensing product readable at all: the install itself did not throw,
# so there is no evidence of failure - report Success but mark it
# unverified so the checklist surfaces it.
Show-Warning "Could not read back the installed key (WMI returned no Windows licensing product)"
Write-Host "========================================" -ForegroundColor White

return (New-ModuleResult -Status "Success" -Message "OEM key installed (read-back failed)" -Verified $false)
