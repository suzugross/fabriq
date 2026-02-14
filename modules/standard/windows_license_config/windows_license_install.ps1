# ========================================
# Windows Product Key Installation Script
# ========================================
# Description: Installs a Windows product key.
# Key source: CSV (license_key.csv) or manual input.
# ========================================

# Check Administrator Privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERROR] This script requires administrator privileges." -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Install Windows Product Key" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
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
    try {
        $allKeys = @(Import-Csv -Path $csvPath -Encoding Default)
        $enabledKeys = @($allKeys | Where-Object { $_.Enabled -eq "1" })

        if ($enabledKeys.Count -gt 0) {
            $productKey = $enabledKeys[0].ProductKey.Trim()
            $keySource = "CSV"
            $keyDesc = $enabledKeys[0].Description

            if ($enabledKeys.Count -gt 1) {
                Write-Host "[WARNING] Multiple enabled keys found. Using first entry." -ForegroundColor Yellow
            }

            Write-Host "[INFO] Product key loaded from CSV" -ForegroundColor Cyan
            Write-Host "  Key:         $productKey" -ForegroundColor White
            if ($keyDesc) {
                Write-Host "  Description: $keyDesc" -ForegroundColor Gray
            }
        }
        else {
            Write-Host "[INFO] No enabled keys in license_key.csv" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "[WARNING] Failed to read license_key.csv: $_" -ForegroundColor Yellow
    }
}
else {
    Write-Host "[INFO] license_key.csv not found" -ForegroundColor Gray
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
    Write-Host "[ERROR] Invalid key format: $productKey" -ForegroundColor Red
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
        Write-Host "[INFO] Uninstalling existing key..." -ForegroundColor Cyan
        try {
            $null = Invoke-CimMethod -InputObject $service -MethodName UninstallProductKey `
                -Arguments @{ProductKeyID = $currentProduct.ID} -ErrorAction Stop
            Write-Host "[SUCCESS] Existing key uninstalled" -ForegroundColor Green
        }
        catch {
            Write-Host "[WARNING] Could not uninstall existing key (will overwrite): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    # Install new key
    Write-Host "[INFO] Installing new product key..." -ForegroundColor Cyan
    $null = Invoke-CimMethod -InputObject $service -MethodName InstallProductKey `
        -Arguments @{ProductKey = $productKey} -ErrorAction Stop
    Write-Host "[SUCCESS] Product key installed" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to install product key: $($_.Exception.Message)" -ForegroundColor Red
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
    Write-Host "  [WARNING] Could not verify installation" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor White

    return (New-ModuleResult -Status "Success" -Message "Key installed (verification pending)")
}
