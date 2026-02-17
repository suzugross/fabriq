# ========================================
# Registry Configuration Script (HKLM)
# ========================================

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
    $items = Import-CsvSafe -Path $csvFile.FullName -Description $csvFile.Name
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
Write-Host "[INFO] Total: $($regItems.Count) Enabled" -NoNewline -ForegroundColor Cyan
if ($skippedCount -gt 0) {
    Write-Host " / $skippedCount Skipped" -NoNewline -ForegroundColor Gray
}
Write-Host "" -ForegroundColor Cyan
Write-Host ""

if ($regItems.Count -eq 0) {
    Show-Info "No valid registry settings found"
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

        Show-Success "Configured"
        $successCount++
    }
    catch {
        Show-Error "$_"
        $failCount++
    }

    Write-Host ""
}

# Summary
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Configuration Results")