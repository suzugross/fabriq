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

# ========================================
# Helper: Query installed partial product keys via /dstatus
# ========================================
# Returns the "Last 5 characters of installed product key" values from
# OSPP.vbs /dstatus as a string array (empty when no products or the
# query fails). OSPP.vbs output literals are English regardless of OS
# locale, so the pattern match is locale-safe.
function Get-InstalledPartialKeys {
    param([Parameter(Mandatory)][string]$OsppPath)
    $partials = @()
    try {
        $statusOutput = & cscript //Nologo "$OsppPath" /dstatus 2>&1 | Out-String
        foreach ($line in ($statusOutput -split "\r?\n")) {
            if ($line -match 'Last 5 characters of installed product key:\s*([A-Za-z0-9]{5})') {
                $partials += $Matches[1].ToUpperInvariant()
            }
        }
    }
    catch { }
    return ,@($partials)
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

    Write-Host "    Key:  $(Get-MaskedKey $item.ProductKey)" -ForegroundColor DarkGray
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
$verifyPass   = 0
$verifyFail   = 0

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
        Show-Skip "Invalid key format (expected XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
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

        if ($exitCode -ne 0) {
            Show-Error "cscript failed (ExitCode=$exitCode): $displayName"
            $failCount++
        }
        else {
            # OSPP.vbs reports /inpkey errors (e.g. 0xC004F050) as text
            # while still exiting 0, so exit code 0 is not proof of
            # registration. Read the state back via /dstatus and require
            # the key's last 5 characters to be present (same
            # distrust-the-exit-code stance as office_license_auth.ps1).
            $keyLast5 = $item.ProductKey.Substring($item.ProductKey.Length - 5).ToUpperInvariant()
            Show-Info "Verifying key registration via /dstatus..."
            $installedKeys = @(Get-InstalledPartialKeys -OsppPath $osppPath)
            if ($installedKeys -notcontains $keyLast5) {
                # License state can lag a moment behind /inpkey; one retry.
                Start-Sleep -Seconds 2
                $installedKeys = @(Get-InstalledPartialKeys -OsppPath $osppPath)
            }

            if ($installedKeys -contains $keyLast5) {
                Write-Host "  [VERIFIED] Key ...$keyLast5 present in /dstatus" -ForegroundColor Green
                Show-Success "Product key registered: $displayName"
                $successCount++
                $verifyPass++
            }
            else {
                Show-Error "Key ...$keyLast5 not present in /dstatus after /inpkey (registration did not take effect): $displayName"
                $failCount++
                $verifyFail++
            }
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
# Verified covers the /dstatus read-back of keys that passed /inpkey:
# $true = every applied key confirmed present, $false = at least one
# missing, $null = nothing reached the read-back (all skipped/failed).
$verified = if (($verifyPass + $verifyFail) -gt 0) { ($verifyFail -eq 0) } else { $null }

return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Office Product Key Installation Results" -Verified $verified)
