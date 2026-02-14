# ========================================
# Printer Driver Uninstallation Script
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Printer Driver Uninstallation" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Step 1: Get Installed Drivers List
# ========================================
Write-Host "[INFO] Scanning for installed printer drivers..." -ForegroundColor Cyan
Write-Host ""

try {
    $allDrivers = Get-PrinterDriver -ErrorAction Stop
}
catch {
    Write-Host "[ERROR] Failed to retrieve printer driver info: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to retrieve printer driver info: $_")
}

# Exclude Microsoft standard drivers
$excludePatterns = @(
    "Microsoft *",
    "Remote Desktop Easy Print",
    "Microsoft enhanced Point and Print compatibility driver",
    "Send to Microsoft OneNote*",
    "Microsoft Shared Fax Driver"
)

$userDrivers = @()
foreach ($drv in $allDrivers) {
    $isExcluded = $false
    foreach ($pattern in $excludePatterns) {
        if ($drv.Name -like $pattern) {
            $isExcluded = $true
            break
        }
    }
    if (-not $isExcluded) {
        $userDrivers += $drv
    }
}

if ($userDrivers.Count -eq 0) {
    Write-Host "[INFO] No uninstallable printer drivers found" -ForegroundColor Yellow
    Write-Host "       (Microsoft standard drivers are excluded)" -ForegroundColor Gray
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No uninstallable printer drivers found")
}

# ========================================
# Step 2: Select Driver
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "[Step 1] Select driver to uninstall" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

for ($i = 0; $i -lt $userDrivers.Count; $i++) {
    $drv = $userDrivers[$i]
    Write-Host "  [$($i + 1)] $($drv.Name)" -ForegroundColor White
    if ($drv.InfPath) {
        Write-Host "        INF: $($drv.InfPath)" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host -NoNewline "Enter number (0 to cancel): "
$choice = Read-Host

if ($choice -eq '0') {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

$choiceNum = 0
if (-not [int]::TryParse($choice, [ref]$choiceNum) -or $choiceNum -lt 1 -or $choiceNum -gt $userDrivers.Count) {
    Write-Host ""
    Write-Host "[ERROR] Invalid number" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid number")
}

$selectedDriver = $userDrivers[$choiceNum - 1]
$driverName = $selectedDriver.Name

Write-Host ""

# ========================================
# Step 3: Check Printers in Use
# ========================================
Write-Host "[INFO] Checking driver usage..." -ForegroundColor Cyan

$usingPrinters = @()
try {
    $allPrinters = Get-Printer -ErrorAction SilentlyContinue
    foreach ($printer in $allPrinters) {
        if ($printer.DriverName -eq $driverName) {
            $usingPrinters += $printer
        }
    }
}
catch {
    # Ignore printer info retrieval failure
}

if ($usingPrinters.Count -gt 0) {
    Write-Host ""
    Write-Host "[WARNING] This driver is currently used by the following printers:" -ForegroundColor Yellow
    foreach ($p in $usingPrinters) {
        Write-Host "  - $($p.Name)" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  These printers will also be deleted" -ForegroundColor Yellow
}

Write-Host ""

# ========================================
# Confirm Uninstallation
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following driver will be uninstalled" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Driver Name: $driverName" -ForegroundColor White
if ($selectedDriver.InfPath) {
    Write-Host "  INF Path:    $($selectedDriver.InfPath)" -ForegroundColor White
}
if ($usingPrinters.Count -gt 0) {
    Write-Host "  In Use:      $($usingPrinters.Count) printers (will be deleted)" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

Write-Host -NoNewline "Do you want to uninstall? (Y/N): "
$confirm = Read-Host

if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Step 4: Delete Printers in Use
# ========================================
if ($usingPrinters.Count -gt 0) {
    Write-Host "[INFO] Deleting printers in use..." -ForegroundColor Cyan

    foreach ($p in $usingPrinters) {
        try {
            Remove-Printer -Name $p.Name -ErrorAction Stop
            Write-Host "[SUCCESS] Deleted printer: $($p.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Failed to delete printer: $($p.Name) - $_" -ForegroundColor Red
            Write-Host ""
            Write-Host "[INFO] Aborting driver uninstallation due to printer deletion failure" -ForegroundColor Yellow
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Failed to delete printer: $($p.Name)")
        }
    }

    Write-Host ""
}

# ========================================
# Step 5: Delete Printer Driver
# ========================================
Write-Host "[INFO] Deleting printer driver: $driverName" -ForegroundColor Cyan

# --- Restart Spooler to release locks ---
Write-Host "[INFO] Restarting Print Spooler to release file locks..." -ForegroundColor Cyan
try {
    Restart-Service -Name "spooler" -Force -ErrorAction Stop
    Start-Sleep -Seconds 3
    Write-Host "[SUCCESS] Print Spooler restarted" -ForegroundColor Green
}
catch {
    Write-Host "[WARNING] Failed to restart Print Spooler: $_" -ForegroundColor Yellow
}
# ----------------------------------------

try {
    Remove-PrinterDriver -Name $driverName -ErrorAction Stop
    Write-Host "[SUCCESS] Deleted printer driver: $driverName" -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to delete printer driver: $_" -ForegroundColor Red
    if ($_.Exception.InnerException) {
         Write-Host "       Detail: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to delete printer driver: $_")
}

Write-Host ""

# ========================================
# Step 6: Remove from Driver Store
# ========================================
Write-Host "[INFO] Checking removal from Driver Store..." -ForegroundColor Cyan

# Search for OEM INF in Driver Store with pnputil
$pnpEnum = & pnputil /enum-drivers 2>&1
$oemInfName = $null

# Find matching OEM INF from pnputil output
$currentOem = $null
foreach ($line in $pnpEnum) {
    $lineStr = "$line".Trim()

    # Extract OEM Name from Published Name line
    if ($lineStr -match '(oem\d+\.inf)') {
        $currentOem = $Matches[1]
    }

    # Check if driver name is included
    if ($currentOem -and $lineStr -match [regex]::Escape($driverName)) {
        $oemInfName = $currentOem
        break
    }

    # Check against original INF path
    if ($currentOem -and $selectedDriver.InfPath -and $lineStr -match [regex]::Escape((Split-Path $selectedDriver.InfPath -Leaf))) {
        $oemInfName = $currentOem
        break
    }

    # Reset on empty line
    if ($lineStr -eq '') {
        $currentOem = $null
    }
}

if ($oemInfName) {
    Write-Host "[INFO] Deleting from Driver Store: $oemInfName" -ForegroundColor Cyan

    $pnpResult = & pnputil /delete-driver $oemInfName /force 2>&1
    $pnpExitCode = $LASTEXITCODE

    if ($pnpExitCode -eq 0) {
        Write-Host "[SUCCESS] Deleted from Driver Store: $oemInfName" -ForegroundColor Green
    }
    else {
        Write-Host "[WARNING] Failed to delete from Driver Store (might be in use by others)" -ForegroundColor Yellow
        foreach ($line in $pnpResult) {
            Write-Host "  $line" -ForegroundColor Gray
        }
    }
}
else {
    Write-Host "[INFO] Could not identify corresponding OEM INF in Driver Store" -ForegroundColor Yellow
    Write-Host "       Please check manually with pnputil /enum-drivers" -ForegroundColor Gray
}

Write-Host ""

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Uninstallation Completed" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Driver: $driverName" -ForegroundColor Green
if ($usingPrinters.Count -gt 0) {
    Write-Host "Printers: $($usingPrinters.Count) deleted" -ForegroundColor Green
}
if ($oemInfName) {
    Write-Host "Store:    $oemInfName deleted" -ForegroundColor Green
}
Write-Host ""

return (New-ModuleResult -Status "Success" -Message "Driver uninstalled: $driverName")