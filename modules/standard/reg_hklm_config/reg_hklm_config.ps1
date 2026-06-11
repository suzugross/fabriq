# ========================================
# Registry Configuration Script (HKLM)
# ========================================

$FORCE_OVERWRITE = $true  # Set to $false to enable idempotency checks and skip already-configured settings

Write-Host ""
Show-Separator
Write-Host "Registry Configuration (HKLM)" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# Find CSV files (matches reg_hklm_list*.csv)
$csvFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter "reg_hklm_list*.csv" -File | Sort-Object Name)

if ($csvFiles.Count -eq 0) {
    Show-Error "No files matching reg_hklm_list*.csv found"
    return (New-ModuleResult -Status "Error" -Message "No files matching reg_hklm_list*.csv found")
}

# Load CSV (Support multiple files)
$allItems = @()
$loadedFileCount = 0

foreach ($csvFile in $csvFiles) {
    $items = Import-ModuleCsv -Path $csvFile.FullName
    if ($null -ne $items) {
        $allItems += $items
        Show-Info "Loaded $($csvFile.Name) ($($items.Count) items)"
        $loadedFileCount++
    }
}

if ($loadedFileCount -eq 0) {
    Show-Error "Failed to load any CSV files"
    return (New-ModuleResult -Status "Error" -Message "Failed to load any CSV files")
}

$regItems = @($allItems | Where-Object { $_.'Enabled' -eq '1' })
$skippedCount = $allItems.Count - $regItems.Count

Write-Host ""
$skipMsg = if ($skippedCount -gt 0) { " ($skippedCount skipped)" } else { "" }
Show-Info "Total: $($regItems.Count) enabled$skipMsg"
Write-Host ""

if ($regItems.Count -eq 0) {
    Show-Info "No valid registry settings found"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid registry settings found")
}

# ========================================
# Registry Numeric Conversion Helpers
# ========================================
# DWORD is unsigned 32-bit on the wire, but .NET stores it as Int32 -
# values 0x80000000-0xFFFFFFFF must be written as the bit-equal negative
# Int32 (a plain [int] cast of '4294967295' simply throws), and they
# read back as negative Int32 so comparisons must normalize the same
# way. Accepts decimal, 0x-prefixed hex, and signed literals from CSV.
function ConvertTo-RegistryDWordValue {
    param([Parameter(Mandatory)][string]$Value)
    $v = $Value.Trim()
    if ($v -match '^-') { return [int]$v }
    $u = if ($v -match '^0[xX]') { [Convert]::ToUInt32($v.Substring(2), 16) } else { [uint32]$v }
    return [BitConverter]::ToInt32([BitConverter]::GetBytes($u), 0)
}

function ConvertTo-RegistryQWordValue {
    param([Parameter(Mandatory)][string]$Value)
    $v = $Value.Trim()
    if ($v -match '^-') { return [long]$v }
    $u = if ($v -match '^0[xX]') { [Convert]::ToUInt64($v.Substring(2), 16) } else { [uint64]$v }
    return [BitConverter]::ToInt64([BitConverter]::GetBytes($u), 0)
}

# ========================================
# Idempotency Helper
# ========================================
function Test-RegistryValueMatch {
    param(
        [string]$Path,
        [string]$Name,
        [string]$ExpectedValue,
        [string]$Type
    )

    try {
        if (-not (Test-Path $Path)) { return $false }

        $prop = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        if ($null -eq $prop) { return $false }

        $currentValue = $prop.$Name

        switch ($Type) {
            'DWord'  { return ([int]$currentValue -eq (ConvertTo-RegistryDWordValue $ExpectedValue)) }
            'QWord'  { return ([long]$currentValue -eq (ConvertTo-RegistryQWordValue $ExpectedValue)) }
            'Binary' {
                $currentHex = ($currentValue | ForEach-Object { '{0:X2}' -f $_ }) -join ''
                $expectedHex = ($ExpectedValue -replace '[^0-9A-Fa-f]', '').ToUpper()
                return ($currentHex -eq $expectedHex)
            }
            'MultiString' {
                $currentJoined = ($currentValue -join "`n")
                return ($currentJoined -eq $ExpectedValue)
            }
            default {
                return ([string]$currentValue -eq [string]$ExpectedValue)
            }
        }
    }
    catch {
        return $false
    }
}

# ========================================
# List Changes
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following registry changes will be applied" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($FORCE_OVERWRITE) {
    Write-Host "[FORCE MODE] All settings will be applied regardless of current state" -ForegroundColor Magenta
    Write-Host ""
}

foreach ($item in $regItems) {
    $checkPath = $item.'KeyPath' -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
    $checkPath = $checkPath -replace '^HKEY_CURRENT_USER', 'HKCU:'
    $isKeyOnly = [string]::IsNullOrWhiteSpace($item.'KeyName')

    if ($isKeyOnly) {
        $isMatch = (Test-Path $checkPath)
    }
    else {
        $checkType = switch ($item.'Type') {
            'REG_SZ' { 'String' }; 'REG_DWORD' { 'DWord' }; 'REG_QWORD' { 'QWord' }
            'REG_BINARY' { 'Binary' }; 'REG_MULTI_SZ' { 'MultiString' }; 'REG_EXPAND_SZ' { 'ExpandString' }
            default { 'String' }
        }
        $isMatch = Test-RegistryValueMatch -Path $checkPath -Name $item.'KeyName' -ExpectedValue $item.'Value' -Type $checkType
    }

    $marker = if ($isMatch) { "[Current]" } else { "[Change]" }
    $markerColor = if ($isMatch) { "Gray" } else { "White" }

    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')  $marker" -ForegroundColor $markerColor
    Write-Host "  Path:  $($item.'KeyPath')"
    if ($isKeyOnly) {
        Write-Host "  Key:   (key only - no value)"
        if ((-not [string]::IsNullOrWhiteSpace($item.'Value')) -or (-not [string]::IsNullOrWhiteSpace($item.'Type'))) {
            Show-Warning "KeyName empty -> key-only mode; supplied Value '$($item.'Value')' (Type '$($item.'Type')') will be ignored"
        }
    }
    else {
        Write-Host "  Key:   $($item.'KeyName')"
        Write-Host "  Type:  $($item.'Type')"
        Write-Host "  Value: $($item.'Value')"
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Confirmation
$cancelResult = Confirm-ModuleExecution -Message "Apply the above registry changes?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""
Show-Info "Starting registry configuration..."
Write-Host ""

# Apply Settings
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($item in $regItems) {
    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')" -ForegroundColor Yellow
    Write-Host "  Path: $($item.'KeyPath')"

    # ----------------------------------------
    # Key-only mode: empty KeyName creates the key with no value
    # (equivalent to: reg add "<key>" /f). FORCE_OVERWRITE does not
    # apply (a value-less key has nothing to overwrite); idempotent.
    # ----------------------------------------
    if ([string]::IsNullOrWhiteSpace($item.'KeyName')) {
        Write-Host "  Key:  (key only - no value)"
        try {
            $keyOnlyPath = $item.'KeyPath'
            $keyOnlyPath = $keyOnlyPath -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
            $keyOnlyPath = $keyOnlyPath -replace '^HKEY_CURRENT_USER', 'HKCU:'
            $keyOnlyPath = $keyOnlyPath -replace '^HKEY_CLASSES_ROOT', 'HKCR:'
            $keyOnlyPath = $keyOnlyPath -replace '^HKEY_USERS', 'HKU:'
            $keyOnlyPath = $keyOnlyPath -replace '^HKEY_CURRENT_CONFIG', 'HKCC:'

            if (Test-Path $keyOnlyPath) {
                Show-Skip "Key already exists"
                $skipCount++
            }
            else {
                Write-Host "  -> Creating registry key" -ForegroundColor Gray
                New-Item -Path $keyOnlyPath -Force | Out-Null
                Show-Success "Key created"
                $successCount++
            }
        }
        catch {
            Show-Error "$_"
            $failCount++
        }
        Write-Host ""
        continue
    }

    Write-Host "  Key:  $($item.'KeyName') = $($item.'Value') ($($item.'Type'))"

    try {
        # Convert path to PowerShell format
        $regPath = $item.'KeyPath'
        $regPath = $regPath -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
        $regPath = $regPath -replace '^HKEY_CURRENT_USER', 'HKCU:'
        $regPath = $regPath -replace '^HKEY_CLASSES_ROOT', 'HKCR:'
        $regPath = $regPath -replace '^HKEY_USERS', 'HKU:'
        $regPath = $regPath -replace '^HKEY_CURRENT_CONFIG', 'HKCC:'

        # Convert type
        $regType = switch ($item.'Type') {
            'REG_SZ'        { 'String' }
            'REG_DWORD'     { 'DWord' }
            'REG_QWORD'     { 'QWord' }
            'REG_BINARY'    { 'Binary' }
            'REG_MULTI_SZ'  { 'MultiString' }
            'REG_EXPAND_SZ' { 'ExpandString' }
            default         { 'String' }
        }

        # Idempotency check: skip if current value matches target
        if (-not $FORCE_OVERWRITE -and (Test-RegistryValueMatch -Path $regPath -Name $item.'KeyName' -ExpectedValue $item.'Value' -Type $regType)) {
            Show-Skip "Already configured"
            $skipCount++
            Write-Host ""
            continue
        }

        # Create key if not exists
        if (-not (Test-Path $regPath)) {
            Write-Host "  -> Creating registry key" -ForegroundColor Gray
            New-Item -Path $regPath -Force | Out-Null
        }

        # Convert value type if needed
        $regValue = $item.'Value'
        if ($regType -eq 'DWord') {
            $regValue = ConvertTo-RegistryDWordValue $regValue
        }
        elseif ($regType -eq 'QWord') {
            $regValue = ConvertTo-RegistryQWordValue $regValue
        }

        # Check existing value
        $existingValue = Get-ItemProperty -Path $regPath -Name $item.'KeyName' -ErrorAction SilentlyContinue

        if ($existingValue) {
            # Update
            Set-ItemProperty -Path $regPath -Name $item.'KeyName' -Value $regValue -Type $regType -Force -ErrorAction Stop
        }
        else {
            # Create
            New-ItemProperty -Path $regPath -Name $item.'KeyName' -Value $regValue -PropertyType $regType -Force -ErrorAction Stop | Out-Null
        }

        Show-Success "Configured"
        $successCount++
    }
    catch {
        Show-Error "$_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying applied settings..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($item in $regItems) {
    $checkPath = $item.'KeyPath' -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
    $checkPath = $checkPath -replace '^HKEY_CURRENT_USER', 'HKCU:'
    $checkPath = $checkPath -replace '^HKEY_CLASSES_ROOT', 'HKCR:'
    $checkPath = $checkPath -replace '^HKEY_USERS', 'HKU:'
    $checkPath = $checkPath -replace '^HKEY_CURRENT_CONFIG', 'HKCC:'
    $checkType = switch ($item.'Type') {
        'REG_SZ'        { 'String' }
        'REG_DWORD'     { 'DWord' }
        'REG_QWORD'     { 'QWord' }
        'REG_BINARY'    { 'Binary' }
        'REG_MULTI_SZ'  { 'MultiString' }
        'REG_EXPAND_SZ' { 'ExpandString' }
        default         { 'String' }
    }
    $displayName = if ($item.'Description') { $item.'Description' } else { $item.'SettingTitle' }

    if ([string]::IsNullOrWhiteSpace($item.'KeyName')) {
        $isMatch = (Test-Path $checkPath)
    }
    else {
        $isMatch = Test-RegistryValueMatch -Path $checkPath -Name $item.'KeyName' -ExpectedValue $item.'Value' -Type $checkType
    }

    if ($isMatch) {
        Write-Host "  [VERIFIED] $displayName" -ForegroundColor Green
        $verifyPass++
    } else {
        Write-Host "  [VERIFY FAILED] $displayName" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)

# Summary
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Configuration Results" -Verified $verified)