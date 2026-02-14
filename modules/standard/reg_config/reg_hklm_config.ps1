# ========================================
# Registry Configuration Script (HKLM)
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Registry Configuration (HKLM)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Find CSV files (matches reg_hklm_list*.csv)
$csvFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter "reg_hklm_list*.csv" -File | Sort-Object Name)

if ($csvFiles.Count -eq 0) {
    Write-Host "[ERROR] No files matching reg_hklm_list*.csv found" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "No files matching reg_hklm_list*.csv found")
}

# Load CSV (Support multiple files)
$allItems = @()
$loadedFileCount = 0

foreach ($csvFile in $csvFiles) {
    try {
        $items = @(Import-Csv -Path $csvFile.FullName -Encoding Default)
        $allItems += $items
        Write-Host "[INFO] Loaded $($csvFile.Name) ($($items.Count) items)" -ForegroundColor Cyan
        $loadedFileCount++
    }
    catch {
        Write-Host "[ERROR] Failed to load $($csvFile.Name): $_" -ForegroundColor Red
    }
}

if ($loadedFileCount -eq 0) {
    Write-Host "[ERROR] Failed to load any CSV files" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load any CSV files")
}

$regItems = @($allItems | Where-Object { $_.'Enabled' -eq '1' })
$skippedCount = $allItems.Count - $regItems.Count

Write-Host ""
Write-Host "[INFO] Total: $($regItems.Count) Enabled" -NoNewline -ForegroundColor Cyan
if ($skippedCount -gt 0) {
    Write-Host " / $skippedCount Skipped" -NoNewline -ForegroundColor Gray
}
Write-Host "" -ForegroundColor Cyan
Write-Host ""

if ($regItems.Count -eq 0) {
    Write-Host "[INFO] No valid registry settings found" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid registry settings found")
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
            'DWord'  { return ([long]$currentValue -eq [long]$ExpectedValue) }
            'QWord'  { return ([long]$currentValue -eq [long]$ExpectedValue) }
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

foreach ($item in $regItems) {
    $checkPath = $item.'KeyPath' -replace '^HKEY_LOCAL_MACHINE', 'HKLM:'
    $checkPath = $checkPath -replace '^HKEY_CURRENT_USER', 'HKCU:'
    $checkType = switch ($item.'Type') {
        'REG_SZ' { 'String' }; 'REG_DWORD' { 'DWord' }; 'REG_QWORD' { 'QWord' }
        'REG_BINARY' { 'Binary' }; 'REG_MULTI_SZ' { 'MultiString' }; 'REG_EXPAND_SZ' { 'ExpandString' }
        default { 'String' }
    }
    $isMatch = Test-RegistryValueMatch -Path $checkPath -Name $item.'KeyName' -ExpectedValue $item.'Value' -Type $checkType

    $marker = if ($isMatch) { "[Current]" } else { "[Change]" }
    $markerColor = if ($isMatch) { "Gray" } else { "White" }

    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')  $marker" -ForegroundColor $markerColor
    Write-Host "  Path:  $($item.'KeyPath')"
    Write-Host "  Key:   $($item.'KeyName')"
    Write-Host "  Type:  $($item.'Type')"
    Write-Host "  Value: $($item.'Value')"
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Confirmation
if (-not (Confirm-Execution -Message "Apply the above registry changes?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""
Write-Host "[INFO] Starting registry configuration..." -ForegroundColor Cyan
Write-Host ""

# Apply Settings
$successCount = 0
$skipCount = 0
$errorCount = 0

foreach ($item in $regItems) {
    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')" -ForegroundColor Yellow
    Write-Host "  Path: $($item.'KeyPath')"
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
        if (Test-RegistryValueMatch -Path $regPath -Name $item.'KeyName' -ExpectedValue $item.'Value' -Type $regType) {
            Write-Host "  [SKIP] Already configured" -ForegroundColor Gray
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
        if ($regType -eq 'DWord' -or $regType -eq 'QWord') {
            $regValue = [int]$regValue
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

        Write-Host "  [SUCCESS] Configured" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "  [ERROR] $_" -ForegroundColor Red
        $errorCount++
    }

    Write-Host ""
}

# Summary
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Configuration Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Success: $successCount items" -ForegroundColor Green
if ($skipCount -gt 0) {
    Write-Host "Skipped: $skipCount items (Already configured)" -ForegroundColor Gray
}
if ($errorCount -gt 0) {
    Write-Host "Failed: $errorCount items" -ForegroundColor Red
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($errorCount -eq 0 -and $successCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    elseif ($successCount -gt 0 -and $skipCount -gt 0) { "Success" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $errorCount")