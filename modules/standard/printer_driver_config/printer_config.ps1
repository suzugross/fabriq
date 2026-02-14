# ========================================
# Printer Registration Script
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Printer Registration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Step 1: Get Printer Info from Environment Variables
# ========================================
Write-Host "[INFO] Loading printer settings..." -ForegroundColor Cyan

$printers = @()

for ($i = 1; $i -le 10; $i++) {
    $name   = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_NAME")
    $driver = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_DRIVER")
    $port   = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_PORT")

    if (-not [string]::IsNullOrEmpty($name)) {
        $printers += [PSCustomObject]@{
            No     = $i
            Name   = $name
            Driver = $driver
            Port   = $port
        }
    }
}

if ($printers.Count -eq 0) {
    Write-Host "[INFO] No printers to register (No printer settings in hostlist)" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No printers to register")
}

# ========================================
# Step 2: Confirm Settings
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following printers will be registered:" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($p in $printers) {
    Write-Host "  [Printer $($p.No)]" -ForegroundColor White
    Write-Host "    Name:       $($p.Name)" -ForegroundColor White
    Write-Host "    Driver:     $($p.Driver)" -ForegroundColor White
    Write-Host "    Port (IP):  $($p.Port)" -ForegroundColor White
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Driver Check
Write-Host "[INFO] Checking driver existence..." -ForegroundColor Cyan

$installedDrivers = @()
try {
    $installedDrivers = Get-PrinterDriver -ErrorAction Stop | Select-Object -ExpandProperty Name
}
catch {
    Write-Host "[WARNING] Failed to list drivers: $_" -ForegroundColor Yellow
}

$missingDrivers = @()
foreach ($p in $printers) {
    if ($installedDrivers.Count -gt 0 -and $p.Driver -notin $installedDrivers) {
        $missingDrivers += $p
    }
}

if ($missingDrivers.Count -gt 0) {
    Write-Host ""
    Write-Host "[WARNING] The following drivers are not installed:" -ForegroundColor Yellow
    foreach ($m in $missingDrivers) {
        Write-Host "  - $($m.Driver) (Printer $($m.No): $($m.Name))" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Please install drivers first via [6] Printer Drivers menu" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host -NoNewline "Do you want to proceed with registration? (Y/N): "
$confirm = Read-Host

if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Step 3: Registration Loop
# ========================================
$successCount = 0
$errorCount = 0

foreach ($p in $printers) {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[Processing] Printer $($p.No): $($p.Name)" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow

    $portName = "IP_$($p.Port)"

    # --- Create Port ---
    Write-Host "[INFO] Creating TCP/IP Port: $portName ($($p.Port))" -ForegroundColor Cyan

    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    if ($existingPort) {
        Write-Host "[INFO] Port already exists: $portName (Skipping)" -ForegroundColor Cyan
    }
    else {
        try {
            Add-PrinterPort -Name $portName -PrinterHostAddress $($p.Port) -ErrorAction Stop
            Write-Host "[SUCCESS] Port created: $portName" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Failed to create port: $portName - $_" -ForegroundColor Red
            $errorCount++
            Write-Host ""
            continue
        }
    }

    # --- Create Printer ---
    Write-Host "[INFO] Creating printer: $($p.Name)" -ForegroundColor Cyan

    $existingPrinter = Get-Printer -Name $p.Name -ErrorAction SilentlyContinue
    if ($existingPrinter) {
        Write-Host "[WARNING] Printer already exists: $($p.Name)" -ForegroundColor Yellow
        Write-Host "[INFO] Deleting and recreating existing printer" -ForegroundColor Cyan

        try {
            Remove-Printer -Name $p.Name -ErrorAction Stop
            Write-Host "[SUCCESS] Deleted existing printer: $($p.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Failed to delete existing printer: $($p.Name) - $_" -ForegroundColor Red
            $errorCount++
            Write-Host ""
            continue
        }
    }

    try {
        Add-Printer -Name $p.Name -DriverName $p.Driver -PortName $portName -ErrorAction Stop
        Write-Host "[SUCCESS] Printer created: $($p.Name)" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to create printer: $($p.Name) - $_" -ForegroundColor Red
        $errorCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Registration Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Success: $successCount items" -ForegroundColor Green
if ($errorCount -gt 0) {
    Write-Host "Failed: $errorCount items" -ForegroundColor Red
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Fail: $errorCount")