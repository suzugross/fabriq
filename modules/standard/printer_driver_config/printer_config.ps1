# ========================================
# Printer Registration Script
# ========================================
# [PURPOSE]
# Register TCP/IP ports and printers from two sources:
#   1. Hostlist environment variables (SELECTED_PRINTER_1..10_*)
#   2. printer_list.csv (module-local, TargetHost filtered)
# Both sources are unioned. Duplicate PrinterName is handled by
# the existing Get-Printer existence check in Step 5.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Printer Registration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Collect Printers from Hostlist Environment Variables
# ========================================
Show-Info "Loading printer settings..."

$printers = @()

for ($i = 1; $i -le 10; $i++) {
    $name   = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_NAME")
    $driver = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_DRIVER")
    $port   = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_PORT")

    if (-not [string]::IsNullOrEmpty($name)) {
        $printers += [PSCustomObject]@{
            Source = "Hostlist"
            Label  = "Printer $i"
            Name   = $name
            Driver = $driver
            Port   = $port
        }
    }
}

# ========================================
# Step 1.5: Collect Printers from printer_list.csv (optional)
# ========================================
# TargetHost column: empty = all hosts, value = exact match with SELECTED_NEW_PCNAME.
# Comparison is case-insensitive.
$csvPath = Join-Path $PSScriptRoot "printer_list.csv"
$currentHost = [Environment]::GetEnvironmentVariable("SELECTED_NEW_PCNAME")

if (Test-Path $csvPath) {
    $csvItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
        -RequiredColumns @("Enabled", "TargetHost", "PrinterName", "DriverName", "PortAddress")

    if ($null -eq $csvItems) {
        Show-Warning "Failed to load printer_list.csv (continuing with hostlist only)"
    }
    else {
        $matched = 0
        foreach ($row in @($csvItems)) {
            $targetHost = if ($null -ne $row.TargetHost) { $row.TargetHost.Trim() } else { "" }
            $isAllHosts = [string]::IsNullOrEmpty($targetHost)
            $isMatch = $isAllHosts -or ($targetHost -ieq $currentHost)

            if (-not $isMatch) { continue }

            $printers += [PSCustomObject]@{
                Source = "CSV"
                Label  = if ($isAllHosts) { "All hosts" } else { "Host: $targetHost" }
                Name   = $row.PrinterName
                Driver = $row.DriverName
                Port   = $row.PortAddress
            }
            $matched++
        }
        if ($matched -gt 0) {
            Show-Info "printer_list.csv: $matched entries matched current host"
        }
    }
}

if ($printers.Count -eq 0) {
    Show-Info "No printers to register (no hostlist entries and no matching CSV entries)"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No printers to register")
}

# ========================================
# Step 2: Display & Confirm Settings
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following printers will be registered:" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$idx = 1
foreach ($p in $printers) {
    $tag = "[{0,-8}]" -f $p.Source
    Write-Host "  $tag $idx. $($p.Label)" -ForegroundColor White
    Write-Host "    Name:       $($p.Name)" -ForegroundColor White
    Write-Host "    Driver:     $($p.Driver)" -ForegroundColor White
    Write-Host "    Port (IP):  $($p.Port)" -ForegroundColor White
    Write-Host ""
    $idx++
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Driver existence check
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
        Write-Host "  - $($m.Driver) ($($m.Source): $($m.Name))" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Please install drivers first via [6] Printer Drivers menu" -ForegroundColor Yellow
    Write-Host ""
}

$cancelResult = Confirm-ModuleExecution -Message "Do you want to proceed with registration?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Registration Loop
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($p in $printers) {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[Processing] $($p.Source): $($p.Name)" -ForegroundColor Yellow
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
# Step 5.5: Post-Apply Verification
# ========================================
# Read back each expected printer and confirm: existence, driver name, port address.
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Post-Apply Verification" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$verifyPass = 0
$verifyFail = 0

foreach ($p in $printers) {
    $expectedPortName = "IP_$($p.Port)"
    $actualPrinter = Get-Printer -Name $p.Name -ErrorAction SilentlyContinue
    $actualPort = Get-PrinterPort -Name $expectedPortName -ErrorAction SilentlyContinue

    $printerOk = $null -ne $actualPrinter
    $driverOk  = $printerOk -and ($actualPrinter.DriverName -eq $p.Driver)
    $portOk    = $null -ne $actualPort -and ($actualPort.PrinterHostAddress -eq $p.Port)
    $bindingOk = $printerOk -and ($actualPrinter.PortName -eq $expectedPortName)

    if ($printerOk -and $driverOk -and $portOk -and $bindingOk) {
        Write-Host "  [VERIFIED] $($p.Name)" -ForegroundColor Green
        $verifyPass++
    }
    else {
        $reason = @()
        if (-not $printerOk) { $reason += "printer not found" }
        if ($printerOk -and -not $driverOk) { $reason += "driver mismatch ($($actualPrinter.DriverName))" }
        if (-not $portOk) { $reason += "port mismatch" }
        if ($printerOk -and -not $bindingOk) { $reason += "binding mismatch ($($actualPrinter.PortName))" }
        Write-Host "  [VERIFY FAILED] $($p.Name) - $($reason -join ', ')" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)

# ========================================
# Step 6: Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Registration Results" -Verified $verified)
