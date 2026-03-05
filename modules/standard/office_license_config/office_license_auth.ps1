# ========================================
# Office License Activation Script
# ========================================
# Activates Office licenses using cscript OSPP.vbs /act.
# Skips if all products are already activated (idempotent).
#
# [NOTES]
# - Requires administrator privileges
# - MAK activation requires internet (TCP 443 / ping fallback check)
# - KMS activation requires KMS server reachability (no internet check)
# - OSPP.vbs path is auto-detected (C2R/MSI, 64/32bit)
# - Evaluates all products: activation runs only when
#   at least one product is not LICENSED
# ========================================

# ========================================
# Helper: Find OSPP.vbs
# ========================================
function Find-OsppVbs {
    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\OSPP.vbs"
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OSPP.vbs"
        "$env:ProgramFiles\Microsoft Office\Office16\OSPP.vbs"
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.vbs"
    )
    foreach ($path in $candidates) {
        if (Test-Path $path) { return $path }
    }
    return $null
}

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "Activate Office License" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Pre-flight
# ========================================

# ---- 1a. Detect OSPP.vbs ----
$osppPath = Find-OsppVbs

if ($null -eq $osppPath) {
    Show-Error "OSPP.vbs not found. Office may not be installed."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "OSPP.vbs not found")
}

Show-Info "Detected OSPP.vbs: $osppPath"

# ---- 1b. Connectivity check (MAK only) ----
# Load office_key.csv to determine if MAK activation is configured.
# If any enabled entry has ActivationType=MAK, verify internet connectivity.
$hasMak = $false
$csvPath = Join-Path $PSScriptRoot "office_key.csv"

if (Test-Path $csvPath) {
    $allKeys = Import-ModuleCsv -Path $csvPath
    if ($null -ne $allKeys) {
        foreach ($key in @($allKeys | Where-Object { $_.Enabled -eq "1" })) {
            if (-not [string]::IsNullOrWhiteSpace($key.ActivationType) -and
                $key.ActivationType -ieq "MAK") {
                $hasMak = $true
                break
            }
        }
    }
}

if ($hasMak) {
    Show-Info "MAK activation detected. Checking internet connectivity..."

    # Primary check: TCP 443 to Microsoft activation server
    $tcpOk = $false
    try {
        $tcpResult = Test-NetConnection -ComputerName "activation.sls.microsoft.com" -Port 443 `
            -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        $tcpOk = $tcpResult.TcpTestSucceeded
    }
    catch {
        $tcpOk = $false
    }

    if ($tcpOk) {
        Show-Success "Activation server reachable (TCP 443)"
    }
    else {
        Show-Warning "TCP 443 to activation server failed. Trying ping fallback..."

        # Fallback check: ICMP ping to 8.8.8.8
        $pingOk = $false
        try {
            $pingOk = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
        }
        catch {
            $pingOk = $false
        }

        if ($pingOk) {
            Show-Warning "Ping to 8.8.8.8 succeeded. Proxy environment suspected. Continuing activation attempt..."
        }
        else {
            Show-Error "No internet connectivity detected. MAK activation requires internet access."
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "No internet connectivity for MAK activation")
        }
    }
}
else {
    Show-Info "KMS activation detected. Skipping internet connectivity check."
}

Write-Host ""


# ========================================
# Step 2: Current License Status Display
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Current License Status" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White

$statusOutput = & cscript //Nologo "$osppPath" /dstatus 2>&1 | Out-String

foreach ($line in ($statusOutput.Trim() -split "\r?\n")) {
    if (-not [string]::IsNullOrWhiteSpace($line)) {
        Write-Host "  $line" -ForegroundColor DarkGray
    }
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""


# ========================================
# Step 3: Idempotency Check (robust multi-product)
# ========================================
# Parse all LICENSE NAME and LICENSE STATUS pairs from /dstatus output.
# Statuses: LICENSED, OOB_GRACE, UNLICENSED, NOTIFICATION, EXTENDED GRACE, etc.
# All products must be LICENSED to skip activation.
# ========================================
$productNames = @()
$productStatuses = @()

foreach ($line in ($statusOutput -split "\r?\n")) {
    if ($line -match '^\s*LICENSE NAME:\s*(.+)$') {
        $productNames += $Matches[1].Trim()
    }
    elseif ($line -match '^\s*LICENSE STATUS:\s*---(.+)---') {
        $productStatuses += $Matches[1].Trim()
    }
}

if ($productStatuses.Count -eq 0) {
    Show-Error "No Office products found in license status output."
    Show-Info "Please install a product key first."
    return (New-ModuleResult -Status "Error" -Message "No Office products detected")
}

# Display per-product status summary
for ($i = 0; $i -lt $productStatuses.Count; $i++) {
    $name = if ($i -lt $productNames.Count) { $productNames[$i] } else { "Product $($i + 1)" }
    $status = $productStatuses[$i]
    $color = if ($status -eq "LICENSED") { "Green" } else { "Yellow" }
    Write-Host "  [$status] $name" -ForegroundColor $color
}

Write-Host ""

# Check if ALL products are already licensed
$unlicensedStatuses = @($productStatuses | Where-Object { $_ -ne "LICENSED" })

if ($unlicensedStatuses.Count -eq 0) {
    Show-Skip "All $($productStatuses.Count) Office product(s) already activated."
    return (New-ModuleResult -Status "Skipped" -Message "All products already activated ($($productStatuses.Count) product(s))")
}

Show-Info "$($unlicensedStatuses.Count) of $($productStatuses.Count) product(s) require activation."


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Activate Office license?"
if ($null -ne $cancelResult) { return $cancelResult }


# ========================================
# Step 5: Trigger Activation
# ========================================
Write-Host ""
Show-Info "Activating Office license..."

$actOutput = & cscript //Nologo "$osppPath" /act 2>&1 | Out-String
$actExitCode = $LASTEXITCODE

foreach ($line in ($actOutput.Trim() -split "\r?\n")) {
    if (-not [string]::IsNullOrWhiteSpace($line)) {
        Write-Host "  $line" -ForegroundColor DarkGray
    }
}

Write-Host ""

if ($actExitCode -ne 0) {
    Show-Warning "cscript /act returned ExitCode=$actExitCode"
}


# ========================================
# Step 6: Verify Result
# ========================================
Start-Sleep -Seconds 3

Show-Info "Verifying activation status..."
$verifyOutput = & cscript //Nologo "$osppPath" /dstatus 2>&1 | Out-String

$verifyNames = @()
$verifyStatuses = @()

foreach ($line in ($verifyOutput -split "\r?\n")) {
    if ($line -match '^\s*LICENSE NAME:\s*(.+)$') {
        $verifyNames += $Matches[1].Trim()
    }
    elseif ($line -match '^\s*LICENSE STATUS:\s*---(.+)---') {
        $verifyStatuses += $Matches[1].Trim()
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor White
Write-Host "Activation Result" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

for ($i = 0; $i -lt $verifyStatuses.Count; $i++) {
    $name = if ($i -lt $verifyNames.Count) { $verifyNames[$i] } else { "Product $($i + 1)" }
    $status = $verifyStatuses[$i]
    $color = switch ($status) {
        "LICENSED"   { "Green" }
        "UNLICENSED" { "Red" }
        default      { "Yellow" }
    }
    Write-Host "  [$status] $name" -ForegroundColor $color
}

Write-Host "========================================" -ForegroundColor White

# Final evaluation
$stillUnlicensed = @($verifyStatuses | Where-Object { $_ -ne "LICENSED" })
$beforeUnlicensedCount = $unlicensedStatuses.Count

if ($verifyStatuses.Count -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "Could not verify activation status")
}
elseif ($stillUnlicensed.Count -eq 0) {
    return (New-ModuleResult -Status "Success" -Message "All $($verifyStatuses.Count) product(s) activated")
}
elseif ($stillUnlicensed.Count -lt $beforeUnlicensedCount) {
    # Some products were newly activated but not all
    $newlyActivated = $beforeUnlicensedCount - $stillUnlicensed.Count
    return (New-ModuleResult -Status "Partial" -Message "$newlyActivated activated, $($stillUnlicensed.Count) still unlicensed")
}
else {
    return (New-ModuleResult -Status "Error" -Message "Activation failed ($($stillUnlicensed.Count) product(s) still unlicensed)")
}
