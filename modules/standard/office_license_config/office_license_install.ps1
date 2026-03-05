# ========================================
# Office Product Key Installation Script
# ========================================
# Registers Office product keys using cscript OSPP.vbs /inpkey.
#
# [NOTES]
# - Requires administrator privileges
# - OSPP.vbs path is auto-detected (C2R/MSI, 64/32bit)
# - OsppPath column in CSV can override auto-detection
# - ProductKey supports ENC: prefix for encrypted values
# - ActivationType (MAK/KMS) is displayed for reference
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
Write-Host "Install Office Product Key" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "office_key.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "ProductKey")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load office_key.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Pre-flight check
# ========================================
$autoOsppPath = Find-OsppVbs

if ($autoOsppPath) {
    Show-Info "Detected OSPP.vbs: $autoOsppPath"
}
else {
    Show-Warning "OSPP.vbs auto-detection failed"
}

# If auto-detect failed, check if any entry has explicit OsppPath
$hasExplicitPath = $false
foreach ($item in $enabledItems) {
    if (-not [string]::IsNullOrWhiteSpace($item.OsppPath)) {
        $hasExplicitPath = $true
        break
    }
}

if ($null -eq $autoOsppPath -and -not $hasExplicitPath) {
    Show-Error "OSPP.vbs not found. Install Office or specify OsppPath in CSV."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "OSPP.vbs not found")
}


# ========================================
# Step 3: Pre-execution display
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Office Product Keys" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { "Office Product Key" }
    $osppPath = if (-not [string]::IsNullOrWhiteSpace($item.OsppPath)) { $item.OsppPath } else { $autoOsppPath }

    # Determine display status
    $encGuard = $null -ne $item.ProductKey -and $item.ProductKey.StartsWith('ENC:')
    $validKey = $item.ProductKey -match '^[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}$'
    $validPath = $null -ne $osppPath -and (Test-Path $osppPath)

    if ($encGuard) {
        Write-Host "  [ENC ERROR] $displayName" -ForegroundColor Red
    }
    elseif ($validKey -and $validPath) {
        Write-Host "  [APPLY] $displayName" -ForegroundColor Yellow
    }
    elseif (-not $validKey) {
        Write-Host "  [INVALID KEY] $displayName" -ForegroundColor Red
    }
    else {
        Write-Host "  [OSPP NOT FOUND] $displayName" -ForegroundColor Red
    }

    Write-Host "    Key:  $($item.ProductKey)" -ForegroundColor DarkGray
    Write-Host "    OSPP: $(if ($osppPath) { $osppPath } else { '(not found)' })" -ForegroundColor DarkGray

    # Activation type display
    $actType = if (-not [string]::IsNullOrWhiteSpace($item.ActivationType)) { $item.ActivationType } else { "(not set)" }
    Write-Host "    Type: $actType" -ForegroundColor DarkGray

    if (-not [string]::IsNullOrWhiteSpace($item.OsppPath)) {
        Write-Host "    Mode: CSV override" -ForegroundColor DarkGray
    }
    else {
        Write-Host "    Mode: Auto-detect" -ForegroundColor DarkGray
    }

    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Install the above product keys?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execution loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { "Office Product Key" }
    $osppPath = if (-not [string]::IsNullOrWhiteSpace($item.OsppPath)) { $item.OsppPath } else { $autoOsppPath }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Guard: ENC: still encrypted
    # ----------------------------------------
    if ($null -ne $item.ProductKey -and $item.ProductKey.StartsWith('ENC:')) {
        Show-Error "ProductKey is still encrypted (ENC:). Decryption may have failed or passphrase was not entered."
        Write-Host ""
        $failCount++
        continue
    }

    # ----------------------------------------
    # Key format validation
    # ----------------------------------------
    if ($item.ProductKey -notmatch '^[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}-[A-Za-z0-9]{5}$') {
        Show-Skip "Invalid key format: $($item.ProductKey)"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # OSPP.vbs existence check
    # ----------------------------------------
    if ($null -eq $osppPath -or -not (Test-Path $osppPath)) {
        Show-Error "OSPP.vbs not found: $(if ($osppPath) { $osppPath } else { '(no path)' })"
        Write-Host ""
        $failCount++
        continue
    }

    # ----------------------------------------
    # Main processing: Register product key
    # ----------------------------------------
    try {
        Show-Info "Registering product key..."

        $output = & cscript //Nologo "$osppPath" "/inpkey:$($item.ProductKey)" 2>&1 | Out-String
        $exitCode = $LASTEXITCODE

        # Display cscript output
        foreach ($line in ($output.Trim() -split "\r?\n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Write-Host "  $line" -ForegroundColor DarkGray
            }
        }

        if ($exitCode -eq 0) {
            Show-Success "Product key registered: $displayName"
            $successCount++
        }
        else {
            Show-Error "cscript failed (ExitCode=$exitCode): $displayName"
            $failCount++
        }
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Office Product Key Installation Results")
