# ========================================
# Registry Configuration (HKCU + Default Profile)
# ========================================

$HIVE_PATH = "$env:SystemDrive\Users\Default\ntuser.dat"
$HIVE_KEY = "HKEY_USERS\Hive"

Write-Host ""
Show-Separator
Write-Host "Registry Config (HKCU + Default Profile)" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# Find CSV files
$csvFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter "reg_hkcu_list*.csv" -File | Sort-Object Name)

if ($csvFiles.Count -eq 0) {
    Show-Error "No files matching reg_hkcu_list*.csv found"
    return (New-ModuleResult -Status "Error" -Message "No files matching reg_hkcu_list*.csv found")
}

# Load CSV
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
Write-Host "  Target 1: HKEY_CURRENT_USER (Current User)" -ForegroundColor Yellow
Write-Host "  Target 2: Default Profile (For New Users)" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $regItems) {
    $checkPath = $item.'KeyPath' -replace '^HKEY_CURRENT_USER', 'HKCU:'
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

# ========================================
# Load Default Profile Hive
# ========================================
$hiveLoaded = $false

if (Test-Path $HIVE_PATH) {
    Show-Info "Loading Default Profile Hive..."
    & reg load $HIVE_KEY $HIVE_PATH 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Show-Success "Hive loaded: $HIVE_KEY"
        $hiveLoaded = $true
        # Create PSDrive for HKU access
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
        }
    }
    else {
        Show-Error "Failed to load Hive"
        Show-Info "Configuring HKCU only"
    }
}
else {
    Show-Error "ntuser.dat not found: $HIVE_PATH"
    Show-Info "Configuring HKCU only"
}
Write-Host ""

# ========================================
# Apply Settings
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($item in $regItems) {
    Write-Host "[$($item.'AdminID')] $($item.'SettingTitle')" -ForegroundColor Yellow

    $regPathOriginal = $item.'KeyPath'
    $regKey = $item.'KeyName'
    $regValue = $item.'Value'

    # Convert Type
    $regType = switch ($item.'Type') {
        'REG_SZ'        { 'String' }
        'REG_DWORD'     { 'DWord' }
        'REG_QWORD'     { 'QWord' }
        'REG_BINARY'    { 'Binary' }
        'REG_MULTI_SZ'  { 'MultiString' }
        'REG_EXPAND_SZ' { 'ExpandString' }
        default         { 'String' }
    }

    # Cast numeric types
    $regValueTyped = $regValue
    if ($regType -eq 'DWord' -or $regType -eq 'QWord') {
        $regValueTyped = [int]$regValue
    }

    $hkcuChanged = $false
    $hiveChanged = $false
    $hasError = $false

    # ========================================
    # 1. Apply to HKCU
    # ========================================
    $hkcuPath = $regPathOriginal -replace '^HKEY_CURRENT_USER', 'HKCU:'

    # Idempotency check for HKCU
    if (Test-RegistryValueMatch -Path $hkcuPath -Name $regKey -ExpectedValue $regValue -Type $regType) {
        Write-Host "  [HKCU] Already configured (Skip)" -ForegroundColor Gray
    }
    else {
        Write-Host "  [HKCU] Applying to current user..." -ForegroundColor Gray

        try {
            if (-not (Test-Path $hkcuPath)) {
                New-Item -Path $hkcuPath -Force | Out-Null
            }

            $existingValue = Get-ItemProperty -Path $hkcuPath -Name $regKey -ErrorAction SilentlyContinue
            if ($existingValue) {
                Set-ItemProperty -Path $hkcuPath -Name $regKey -Value $regValueTyped -Type $regType -Force -ErrorAction Stop
            }
            else {
                New-ItemProperty -Path $hkcuPath -Name $regKey -Value $regValueTyped -PropertyType $regType -Force -ErrorAction Stop | Out-Null
            }

            Write-Host "  [HKCU] Configured" -ForegroundColor Green
            $hkcuChanged = $true
        }
        catch {
            Write-Host "  [HKCU] Error: $_" -ForegroundColor Red
            $failCount++
            $hasError = $true
            Write-Host ""
            continue
        }
    }

    # ========================================
    # 2. Apply to Default Profile (HIVE)
    # ========================================
    if ($hiveLoaded) {
        $hivePsPath = $regPathOriginal -replace '^HKEY_CURRENT_USER', 'HKU:\Hive'

        # Idempotency check for HIVE
        if (Test-RegistryValueMatch -Path $hivePsPath -Name $regKey -ExpectedValue $regValue -Type $regType) {
            Write-Host "  [HIVE] Already configured (Skip)" -ForegroundColor Gray
        }
        else {
            Write-Host "  [HIVE] Applying to Default Profile..." -ForegroundColor Gray

            try {
                if (-not (Test-Path $hivePsPath)) {
                    New-Item -Path $hivePsPath -Force | Out-Null
                }

                $existingHiveValue = Get-ItemProperty -Path $hivePsPath -Name $regKey -ErrorAction SilentlyContinue
                if ($existingHiveValue) {
                    Set-ItemProperty -Path $hivePsPath -Name $regKey -Value $regValueTyped -Type $regType -Force -ErrorAction Stop
                }
                else {
                    New-ItemProperty -Path $hivePsPath -Name $regKey -Value $regValueTyped -PropertyType $regType -Force -ErrorAction Stop | Out-Null
                }

                Write-Host "  [HIVE] Configured" -ForegroundColor Green
                $hiveChanged = $true
            }
            catch {
                Write-Host "  [HIVE] Error: $_" -ForegroundColor Red
                $failCount++
                $hasError = $true
                Write-Host ""
                continue
            }
        }
    }
    else {
        Write-Host "  [HIVE] Skipped (Hive not loaded)" -ForegroundColor Gray
    }

    if (-not $hasError) {
        if (-not $hkcuChanged -and -not $hiveChanged) {
            $skipCount++
        }
        else {
            $successCount++
        }
    }
    Write-Host ""
}

# ========================================
# Unload Hive
# ========================================
if ($hiveLoaded) {
    Show-Info "Unloading Hive..."

    if (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name HKU -Force
    }

    [gc]::Collect()
    Start-Sleep -Seconds 1

    & reg unload $HIVE_KEY 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Show-Success "Hive unloaded"
    }
    else {
        Show-Warning "Failed to unload Hive. Retrying..."
        Start-Sleep -Seconds 2
        [gc]::Collect()
        & reg unload $HIVE_KEY 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Show-Success "Hive unloaded (Retry success)"
        }
        else {
            Show-Error "Failed to unload Hive. Please unload manually."
        }
    }
    Write-Host ""
}

# Summary
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Configuration Results")