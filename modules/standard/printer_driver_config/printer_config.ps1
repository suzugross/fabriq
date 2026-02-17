# ========================================
# Printer Registration Script
# ========================================

Write-Host ""
Show-Separator
Write-Host "Printer Registration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Get Printer Info from Environment Variables
# ========================================
Show-Info "Loading printer settings..."

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
    Show-Info "No printers to register (No printer settings in hostlist)"
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
Show-Info "Checking driver existence..."

$installedDrivers = @()
try {
    $installedDrivers = Get-PrinterDriver -ErrorAction Stop | Select-Object -ExpandProperty Name
}
catch {
    Show-Warning "Failed to list drivers: $_"
}

$missingDrivers = @()
foreach ($p in $printers) {
    if ($installedDrivers.Count -gt 0 -and $p.Driver -notin $installedDrivers) {
        $missingDrivers += $p
    }
}

if ($missingDrivers.Count -gt 0) {
    Write-Host ""
    Show-Warning "The following drivers are not installed:"
    foreach ($m in $missingDrivers) {
        Write-Host "  - $($m.Driver) (Printer $($m.No): $($m.Name))" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Please install drivers first via [6] Printer Drivers menu" -ForegroundColor Yellow
    Write-Host ""
}

$cancelResult = Confirm-ModuleExecution -Message "Do you want to proceed with registration?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 3: Registration Loop
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($p in $printers) {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[Processing] Printer $($p.No): $($p.Name)" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow

    $portName = "IP_$($p.Port)"

    # --- Create Port ---
    Show-Info "Creating TCP/IP Port: $portName ($($p.Port))"

    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    if ($existingPort) {
        Show-Info "Port already exists: $portName (Skipping)"
    }
    else {
        try {
            Add-PrinterPort -Name $portName -PrinterHostAddress $($p.Port) -ErrorAction Stop
            Show-Success "Port created: $portName"
        }
        catch {
            Show-Error "Failed to create port: $portName - $_"
            $failCount++
            Write-Host ""
            continue
        }
    }

    # --- Create Printer ---
    Show-Info "Creating printer: $($p.Name)"

    $existingPrinter = Get-Printer -Name $p.Name -ErrorAction SilentlyContinue
    if ($existingPrinter) {
        Show-Skip "Printer already exists: $($p.Name)"
        $skipCount++
        Write-Host ""
        continue
    }

    try {
        Add-Printer -Name $p.Name -DriverName $p.Driver -PortName $portName -ErrorAction Stop
        Show-Success "Printer created: $($p.Name)"
        $successCount++
    }
    catch {
        Show-Error "Failed to create printer: $($p.Name) - $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Registration Results")