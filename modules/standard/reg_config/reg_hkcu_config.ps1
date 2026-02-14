# ========================================
# Registry Configuration (HKCU + Default Profile)
# ========================================

$HIVE_PATH = "$env:SystemDrive\Users\Default\ntuser.dat"
$HIVE_KEY = "HKEY_USERS\Hive"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Registry Config (HKCU + Default Profile)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Find CSV files
$csvFiles = @(Get-ChildItem -Path $PSScriptRoot -Filter "reg_hkcu_list*.csv" -File | Sort-Object Name)

if ($csvFiles.Count -eq 0) {
    Write-Host "[ERROR] No files matching reg_hkcu_list*.csv found" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "No files matching reg_hkcu_list*.csv found")
}

# Load CSV
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
if (-not (Confirm-Execution -Message "Apply the above registry changes?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""
Write-Host "[INFO] Starting registry configuration..." -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load Default Profile Hive
# ========================================
$hiveLoaded = $false

if (Test-Path $HIVE_PATH) {
    Write-Host "[INFO] Loading Default Profile Hive..." -ForegroundColor Cyan
    & reg load $HIVE_KEY $HIVE_PATH 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[SUCCESS] Hive loaded: $HIVE_KEY" -ForegroundColor Green
        $hiveLoaded = $true
        # Create PSDrive for HKU access
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
        }
    }
    else {
        Write-Host "[ERROR] Failed to load Hive" -ForegroundColor Red
        Write-Host "[INFO] Configuring HKCU only" -ForegroundColor Yellow
    }
}
else {
    Write-Host "[ERROR] ntuser.dat not found: $HIVE_PATH" -ForegroundColor Red
    Write-Host "[INFO] Configuring HKCU only" -ForegroundColor Yellow
}
Write-Host ""

# ========================================
# Apply Settings
# ========================================
$successCount = 0
$skipCount = 0
$errorCount = 0

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
            $errorCount++
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
                $errorCount++
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
    Write-Host "[INFO] Unloading Hive..." -ForegroundColor Cyan

    if (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue) {
        Remove-PSDrive -Name HKU -Force
    }

    [gc]::Collect()
    Start-Sleep -Seconds 1

    & reg unload $HIVE_KEY 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[SUCCESS] Hive unloaded" -ForegroundColor Green
    }
    else {
        Write-Host "[ERROR] Failed to unload Hive. Retrying..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        [gc]::Collect()
        & reg unload $HIVE_KEY 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[SUCCESS] Hive unloaded (Retry success)" -ForegroundColor Green
        }
        else {
            Write-Host "[ERROR] Failed to unload Hive. Please unload manually." -ForegroundColor Red
        }
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