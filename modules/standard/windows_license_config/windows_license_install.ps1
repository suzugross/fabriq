# ========================================
# Windows Product Key Installation Script
# ========================================
# Description: Installs a Windows product key.
# Key source: CSV (license_key.csv) or manual input.
# ========================================

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
    $allKeys = Import-CsvSafe -Path $csvPath -Description "license_key.csv"
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
            Write-Host "  Key:         $productKey" -ForegroundColor White
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
    Show-Error "Invalid key format: $productKey"
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
Write-Host "New Key:    $productKey (Source: $keySource)" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: Confirm
# ========================================
if (-not (Confirm-Execution -Message "Install this product key?")) {
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

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
# Step 6: Verify Installation
# ========================================
Start-Sleep -Seconds 1

$finalCheck = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue |
              Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } |
              Select-Object -First 1

Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host "Installation Result" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

if ($finalCheck) {
    $editionName = $finalCheck.Name.Split(',')[0]
    Write-Host "  Edition:        $editionName" -ForegroundColor Green
    Write-Host "  Partial Key:    $($finalCheck.PartialProductKey)" -ForegroundColor Green
    Write-Host "  License Status: $(Get-LicenseStatusText $finalCheck.LicenseStatus)"
    Write-Host "========================================" -ForegroundColor White

    return (New-ModuleResult -Status "Success" -Message "Key installed (Partial: $($finalCheck.PartialProductKey))")
}
else {
    Show-Warning "Could not verify installation"
    Write-Host "========================================" -ForegroundColor White

    return (New-ModuleResult -Status "Success" -Message "Key installed (verification pending)")
}
