# ========================================
# Windows Product Key Installation Script
# ========================================
# Description: Installs a Windows product key.
# Key source: CSV (license_key.csv) or manual input.
# ========================================

# ========================================
# Helper: Mask product key for safe display
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

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "  Install Windows Product Key" -ForegroundColor Cyan
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
# Step 1: Get Product Key (CSV -> Manual)
# ========================================
$productKey = $null
$keySource = ""

$csvPath = Join-Path $PSScriptRoot "license_key.csv"
if (Test-Path $csvPath) {
    $allKeys = Import-ModuleCsv -Path $csvPath
    if ($null -ne $allKeys -and $allKeys.Count -gt 0) {
        $enabledKeys = @($allKeys | Where-Object { $_.Enabled -eq "1" })

        if ($enabledKeys.Count -gt 0) {
            $productKey = $enabledKeys[0].ProductKey.Trim()
            $keySource = "CSV"
            $keyDesc = $enabledKeys[0].Description

            if ($enabledKeys.Count -gt 1) {
                Show-Warning "Multiple enabled keys found. Using first entry."
            }

            Show-Info "Product key loaded from CSV"
            Write-Host "  Key:         $(Get-MaskedKey $productKey)" -ForegroundColor White
            if ($keyDesc) {
                Write-Host "  Description: $keyDesc" -ForegroundColor Gray
            }
        }
        else {
            Show-Info "No enabled keys in license_key.csv"
        }
    }
}
else {
    Show-Info "license_key.csv not found (manual input mode)"
}

# Manual input fallback
if ([string]::IsNullOrWhiteSpace($productKey)) {
    Write-Host ""
    Write-Host "Enter product key manually (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)" -ForegroundColor Yellow
    Write-Host -NoNewline "Product Key: "
    $productKey = (Read-Host).Trim()
    $keySource = "Manual"

    if ([string]::IsNullOrWhiteSpace($productKey)) {
        return (New-ModuleResult -Status "Cancelled" -Message "No product key provided")
    }
}

# ========================================
# Step 2: Validate Key Format
# ========================================
if ($productKey -notmatch '^[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}$') {
    Show-Error "Invalid key format (expected XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
    return (New-ModuleResult -Status "Error" -Message "Invalid product key format")
}

# ========================================
# Step 3: Show Current License Status
# ========================================
Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Current License Status" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

$currentProduct = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
                  Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } |
                  Select-Object -First 1

if ($currentProduct) {
    $editionName = $currentProduct.Name.Split(',')[0]
    Write-Host "  Edition:        $editionName"
    Write-Host "  Partial Key:    $($currentProduct.PartialProductKey)"
    Write-Host "  License Status: $(Get-LicenseStatusText $currentProduct.LicenseStatus)"
}
else {
    Write-Host "  No existing product key found" -ForegroundColor Gray
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""
Write-Host "New Key:    $(Get-MaskedKey $productKey) (Source: $keySource)" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: Confirm
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Install this product key?"
if ($null -ne $cancelResult) { return $cancelResult }

# ========================================
# Step 5: Install Product Key
# ========================================
Write-Host ""
try {
    $service = Get-CimInstance -ClassName SoftwareLicensingService

    # Uninstall existing key if present (optional, InstallProductKey overwrites anyway)
    if ($currentProduct) {
        Show-Info "Uninstalling existing key..."
        try {
            $null = Invoke-CimMethod -InputObject $service -MethodName UninstallProductKey `
                -Arguments @{ProductKeyID = $currentProduct.ID} -ErrorAction Stop
            Show-Success "Existing key uninstalled"
        }
        catch {
            Show-Warning "Could not uninstall existing key (will overwrite): $($_.Exception.Message)"
        }
    }

    # Install new key
    Show-Info "Installing new product key..."
    $null = Invoke-CimMethod -InputObject $service -MethodName InstallProductKey `
        -Arguments @{ProductKey = $productKey} -ErrorAction Stop
    Show-Success "Product key installed"
}
catch {
    Show-Error "Failed to install product key: $($_.Exception.Message)"
    return (New-ModuleResult -Status "Error" -Message "Install failed: $($_.Exception.Message)")
}

# ========================================
# Step 6: Verify Installation (read-back against the entered key)
# ========================================
# Mere existence of a Windows product with a PartialProductKey is not
# verification - a leftover OLD key looks identical. Match the last 5
# characters of the entered key (the part Get-MaskedKey keeps visible
# by design, so no new exposure). Retry to absorb WMI refresh latency.
$expectedPartial = $productKey.Substring($productKey.Length - 5).ToUpper()
$matchedProduct  = $null
$winProducts     = @()

for ($attempt = 1; $attempt -le 3; $attempt++) {
    Start-Sleep -Seconds 2
    $winProducts = @(Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
                     Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" })
    $matchedProduct = $winProducts |
        Where-Object { "$($_.PartialProductKey)".ToUpper() -eq $expectedPartial } |
        Select-Object -First 1
    if ($matchedProduct) { break }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host "Installation Result" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

if ($matchedProduct) {
    $editionName = $matchedProduct.Name.Split(',')[0]
    Write-Host "  Edition:        $editionName" -ForegroundColor Green
    Write-Host "  Partial Key:    $($matchedProduct.PartialProductKey) (matches entered key)" -ForegroundColor Green
    Write-Host "  License Status: $(Get-LicenseStatusText $matchedProduct.LicenseStatus)"
    Write-Host "========================================" -ForegroundColor White

    return (New-ModuleResult -Status "Success" -Message "Key installed (Partial: $($matchedProduct.PartialProductKey))" -Verified $true)
}

if ($winProducts.Count -gt 0) {
    # Products are readable but none carries the entered key - the OLD
    # key is still in effect. Definite contradiction -> fail closed.
    $actualPartials = ($winProducts | ForEach-Object { $_.PartialProductKey }) -join ', '
    Show-Error "Installed key mismatch (expected last5: $expectedPartial, actual: $actualPartials)"
    Write-Host "========================================" -ForegroundColor White

    return (New-ModuleResult -Status "Error" -Message "Key mismatch after install (expected: $expectedPartial, actual: $actualPartials)")
}

# No licensing product readable at all: the install itself did not throw,
# so there is no evidence of failure - report Success but mark it
# unverified so the checklist surfaces it.
Show-Warning "Could not read back the installed key (WMI returned no Windows licensing product)"
Write-Host "========================================" -ForegroundColor White

return (New-ModuleResult -Status "Success" -Message "Key installed (read-back failed)" -Verified $false)
