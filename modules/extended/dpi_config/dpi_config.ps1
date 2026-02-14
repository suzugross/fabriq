# ========================================
# Display DPI Scaling Configuration Script
# ========================================
# Modifies per-monitor DPI scaling via registry
# (HKCU + Default User Hive) based on CSV
# configuration (dpi_list.csv).
# Sign-out or restart required for changes.
# ========================================

# Check Administrator Privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Administrator privileges are required."
    Write-Warning "Please run PowerShell as Administrator and try again."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Display DPI Scaling Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Constants
# ========================================
$script:PerMonitorBasePath = 'HKCU:\Control Panel\Desktop\PerMonitorSettings'
$script:GraphicsConfigPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration'
$HIVE_PATH = "$env:SystemDrive\Users\Default\ntuser.dat"
$HIVE_KEY = "HKEY_USERS\Hive"

# DPI Value Mapping: Percent -> Registry DWord (Int32)
$script:DpiValueMap = @{
    100 = [int]-2
    125 = [int]-1
    150 = [int]0
    175 = [int]1
    200 = [int]2
}

# Reverse mapping: DWord -> Percent string
$script:DpiPercentMap = @{
    -2 = "100%"
    -1 = "125%"
    0  = "150%"
    1  = "175%"
    2  = "200%"
}

# ========================================
# Helper Functions
# ========================================

function Convert-DpiToPercent {
    param([int]$DpiValue)
    if ($script:DpiPercentMap.ContainsKey($DpiValue)) {
        return $script:DpiPercentMap[$DpiValue]
    }
    return "Unknown($DpiValue)"
}

function Get-CurrentDpiValue {
    param([string]$KeyPath)
    try {
        $props = Get-ItemProperty $KeyPath -Name 'DpiValue' -ErrorAction Stop
        return [int]$props.DpiValue
    }
    catch {
        return $null
    }
}

function Find-PerMonitorKeys {
    param(
        [string]$HardwareID,
        [string]$BasePath = $script:PerMonitorBasePath
    )
    if (-not (Test-Path $BasePath)) { return @() }
    $allKeys = Get-ChildItem $BasePath -ErrorAction SilentlyContinue
    if (-not $allKeys) { return @() }

    if ([string]::IsNullOrWhiteSpace($HardwareID)) {
        return @($allKeys)
    }
    return @($allKeys | Where-Object { $_.PSChildName -like "$HardwareID*" })
}

function Find-DisplayKeyNames {
    param([string]$HardwareID)
    if (-not (Test-Path $script:GraphicsConfigPath)) { return @() }
    $allKeys = Get-ChildItem $script:GraphicsConfigPath -ErrorAction SilentlyContinue
    if (-not $allKeys) { return @() }

    if ([string]::IsNullOrWhiteSpace($HardwareID)) {
        return @($allKeys | ForEach-Object { $_.PSChildName })
    }
    return @($allKeys | Where-Object { $_.PSChildName -like "$HardwareID*" } |
             ForEach-Object { $_.PSChildName })
}

function Select-DisplayInteractive {
    param([string]$Description)

    # Collect displays from PerMonitorSettings + GraphicsDrivers
    $pmKeys = Find-PerMonitorKeys -HardwareID ""
    $gdKeyNames = Find-DisplayKeyNames -HardwareID ""

    $allDisplays = @()
    $seenNames = @{}

    foreach ($k in $pmKeys) {
        $path = Join-Path $script:PerMonitorBasePath $k.PSChildName
        $val = Get-CurrentDpiValue -KeyPath $path
        $currentStr = if ($null -ne $val) { Convert-DpiToPercent $val } else { "Not set" }
        $allDisplays += [PSCustomObject]@{
            Name       = $k.PSChildName
            Source     = "PerMonitorSettings"
            CurrentDpi = $currentStr
        }
        $seenNames[$k.PSChildName] = $true
    }

    foreach ($name in $gdKeyNames) {
        if (-not $seenNames.ContainsKey($name)) {
            $allDisplays += [PSCustomObject]@{
                Name       = $name
                Source     = "GraphicsDrivers"
                CurrentDpi = "Not set"
            }
        }
    }

    if ($allDisplays.Count -eq 0) {
        Write-Host "[ERROR] No display keys found in registry" -ForegroundColor Red
        return $null
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Available Displays" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    $idx = 0
    foreach ($d in $allDisplays) {
        $idx++
        Write-Host "[$idx] $($d.Name)" -ForegroundColor White
        Write-Host "    Current DPI: $($d.CurrentDpi) ($($d.Source))" -ForegroundColor Gray
        Write-Host ""
    }

    if (-not [string]::IsNullOrEmpty($Description)) {
        Write-Host "Description: $Description" -ForegroundColor Cyan
    }
    Write-Host ""

    $selection = Read-Host "Select display number (or 0 to skip)"
    $selNum = 0
    if (-not [int]::TryParse($selection, [ref]$selNum)) {
        Write-Host "[INFO] Invalid input, skipping" -ForegroundColor Yellow
        return $null
    }

    if ($selNum -le 0 -or $selNum -gt $allDisplays.Count) {
        Write-Host "[INFO] Skipped" -ForegroundColor Yellow
        return $null
    }

    return $allDisplays[$selNum - 1].Name
}

function Select-ScaleInteractive {
    Write-Host ""
    Write-Host "Available Scale Settings:" -ForegroundColor Cyan
    Write-Host "  [1] 100%" -ForegroundColor White
    Write-Host "  [2] 125%" -ForegroundColor White
    Write-Host "  [3] 150% (Recommended)" -ForegroundColor Green
    Write-Host "  [4] 175%" -ForegroundColor White
    Write-Host "  [5] 200%" -ForegroundColor White
    Write-Host ""

    $selection = Read-Host "Select scale (or 0 to skip)"
    $selNum = 0
    if (-not [int]::TryParse($selection, [ref]$selNum)) {
        Write-Host "[INFO] Invalid input, skipping" -ForegroundColor Yellow
        return $null
    }

    $scaleOptions = @(100, 125, 150, 175, 200)
    if ($selNum -le 0 -or $selNum -gt $scaleOptions.Count) {
        Write-Host "[INFO] Skipped" -ForegroundColor Yellow
        return $null
    }

    return $scaleOptions[$selNum - 1]
}

function Write-DpiValue {
    param(
        [string]$FullKeyName,
        [int]$DpiValue,
        [bool]$HiveLoaded
    )

    $hkcuChanged = $false
    $hiveChanged = $false
    $hasError = $false

    # --- HKCU ---
    $hkcuPath = Join-Path $script:PerMonitorBasePath $FullKeyName
    $currentHkcu = Get-CurrentDpiValue -KeyPath $hkcuPath

    if ($null -ne $currentHkcu -and $currentHkcu -eq $DpiValue) {
        Write-Host "  [HKCU] Already configured (Skip)" -ForegroundColor Gray
    }
    else {
        try {
            if (-not (Test-Path $hkcuPath)) {
                $null = New-Item -Path $hkcuPath -Force
            }
            $existing = Get-ItemProperty $hkcuPath -Name 'DpiValue' -ErrorAction SilentlyContinue
            if ($existing) {
                Set-ItemProperty -Path $hkcuPath -Name 'DpiValue' -Value $DpiValue -Type DWord -Force -ErrorAction Stop
            }
            else {
                $null = New-ItemProperty -Path $hkcuPath -Name 'DpiValue' -Value $DpiValue -PropertyType DWord -Force -ErrorAction Stop
            }
            Write-Host "  [HKCU] Configured" -ForegroundColor Green
            $hkcuChanged = $true
        }
        catch {
            Write-Host "  [HKCU] Error: $_" -ForegroundColor Red
            $hasError = $true
        }
    }

    # --- HIVE (Default User) ---
    if ($HiveLoaded) {
        $hivePath = "HKU:\Hive\Control Panel\Desktop\PerMonitorSettings\$FullKeyName"
        $currentHive = Get-CurrentDpiValue -KeyPath $hivePath

        if ($null -ne $currentHive -and $currentHive -eq $DpiValue) {
            Write-Host "  [HIVE] Already configured (Skip)" -ForegroundColor Gray
        }
        else {
            try {
                if (-not (Test-Path $hivePath)) {
                    $null = New-Item -Path $hivePath -Force
                }
                $existingHive = Get-ItemProperty $hivePath -Name 'DpiValue' -ErrorAction SilentlyContinue
                if ($existingHive) {
                    Set-ItemProperty -Path $hivePath -Name 'DpiValue' -Value $DpiValue -Type DWord -Force -ErrorAction Stop
                }
                else {
                    $null = New-ItemProperty -Path $hivePath -Name 'DpiValue' -Value $DpiValue -PropertyType DWord -Force -ErrorAction Stop
                }
                Write-Host "  [HIVE] Configured" -ForegroundColor Green
                $hiveChanged = $true
            }
            catch {
                Write-Host "  [HIVE] Error: $_" -ForegroundColor Red
                $hasError = $true
            }
        }
    }
    else {
        Write-Host "  [HIVE] Skipped (Hive not loaded)" -ForegroundColor Gray
    }

    return [PSCustomObject]@{
        HkcuChanged = $hkcuChanged
        HiveChanged = $hiveChanged
        HasError    = $hasError
    }
}

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "dpi_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] dpi_list.csv not found: $csvPath" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "dpi_list.csv not found")
}

try {
    $allItems = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load dpi_list.csv: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to load dpi_list.csv: $_")
}

if ($allItems.Count -eq 0) {
    Write-Host "[ERROR] dpi_list.csv contains no data" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "dpi_list.csv contains no data")
}

# Filter enabled entries
$items = @($allItems | Where-Object { $_.Enabled -eq "1" })

if ($items.Count -eq 0) {
    Write-Host "[INFO] No enabled entries in dpi_list.csv" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

Write-Host "[INFO] Loaded $($items.Count) enabled entries (total: $($allItems.Count))" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Validate CSV Data
# ========================================
$validItems = @()
foreach ($item in $items) {
    $sp = $item.ScalePercent.Trim()

    # ScalePercent empty = interactive selection (valid)
    if ([string]::IsNullOrWhiteSpace($sp)) {
        $validItems += $item
        continue
    }

    $spNum = 0
    if (-not [int]::TryParse($sp, [ref]$spNum)) {
        Write-Host "[WARN] Invalid ScalePercent for '$($item.Description)': '$sp' — skipping" -ForegroundColor Yellow
        continue
    }
    if (-not $script:DpiValueMap.ContainsKey($spNum)) {
        Write-Host "[WARN] Unsupported ScalePercent '$spNum' for '$($item.Description)' (valid: 100,125,150,175,200) — skipping" -ForegroundColor Yellow
        continue
    }
    $validItems += $item
}

if ($validItems.Count -eq 0) {
    Write-Host "[ERROR] No valid entries after validation" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No valid entries after validation")
}

# ========================================
# Resolve Display Targets
# ========================================
$targets = @()

foreach ($item in $validItems) {
    $hwId = $item.HardwareID.Trim()
    $sp = $item.ScalePercent.Trim()

    $interactiveDisplay = [string]::IsNullOrWhiteSpace($hwId)
    $interactiveScale = [string]::IsNullOrWhiteSpace($sp)

    if ($interactiveDisplay) {
        $targets += [PSCustomObject]@{
            Item               = $item
            MatchedKeyNames    = @()
            InteractiveDisplay = $true
            InteractiveScale   = $interactiveScale
        }
    }
    else {
        # Search PerMonitorSettings first
        $pmMatched = Find-PerMonitorKeys -HardwareID $hwId
        $matchedNames = @($pmMatched | ForEach-Object { $_.PSChildName })

        # Fallback: search GraphicsDrivers\Configuration
        if ($matchedNames.Count -eq 0) {
            $matchedNames = Find-DisplayKeyNames -HardwareID $hwId
        }

        $targets += [PSCustomObject]@{
            Item               = $item
            MatchedKeyNames    = $matchedNames
            InteractiveDisplay = $false
            InteractiveScale   = $interactiveScale
        }
    }
}

# ========================================
# List Settings with Idempotency Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Target DPI Scaling Settings" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$index = 0
foreach ($target in $targets) {
    $index++
    $item = $target.Item
    $sp = $item.ScalePercent.Trim()
    $scaleStr = if ($target.InteractiveScale) { "TBD (Interactive)" } else { "${sp}%" }

    if ($target.InteractiveDisplay) {
        Write-Host "[$index] (Interactive Selection) -> $scaleStr" -ForegroundColor Yellow
        Write-Host "    Display and/or scale will be selected during apply phase" -ForegroundColor Gray
        Write-Host "    $($item.Description)" -ForegroundColor Gray
        Write-Host ""
        continue
    }

    if ($target.MatchedKeyNames.Count -eq 0) {
        Write-Host "[$index] $($item.HardwareID) -> $scaleStr  [ERROR]" -ForegroundColor Red
        Write-Host "    No display found matching '$($item.HardwareID)'" -ForegroundColor Red
        Write-Host "    $($item.Description)" -ForegroundColor Gray
        Write-Host ""
        continue
    }

    foreach ($keyName in $target.MatchedKeyNames) {
        $hkcuPath = Join-Path $script:PerMonitorBasePath $keyName
        $currentVal = Get-CurrentDpiValue -KeyPath $hkcuPath
        $currentStr = if ($null -ne $currentVal) { Convert-DpiToPercent $currentVal } else { "Not set" }

        if (-not $target.InteractiveScale) {
            $targetDpi = $script:DpiValueMap[[int]$sp]
            if ($null -ne $currentVal -and $currentVal -eq $targetDpi) {
                $marker = "[Current]"
                $markerColor = "Gray"
            }
            else {
                $marker = "[Change]"
                $markerColor = "White"
            }
        }
        else {
            $marker = "[Interactive]"
            $markerColor = "Yellow"
        }

        Write-Host "[$index] $($item.HardwareID) -> $scaleStr  $marker" -ForegroundColor $markerColor
        Write-Host "    Matched: $keyName" -ForegroundColor Gray
        Write-Host "    Current: $currentStr | $($item.Description)" -ForegroundColor Gray
        Write-Host ""
    }
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Apply the above DPI scaling settings?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

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
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            $null = New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS
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
Write-Host "--- Applying DPI Scaling Settings ---" -ForegroundColor Cyan
Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0

$index = 0
foreach ($target in $targets) {
    $index++
    $item = $target.Item
    $sp = $item.ScalePercent.Trim()

    # Resolve display key name(s)
    $keyNamesToProcess = @()
    if ($target.InteractiveDisplay) {
        $selectedName = Select-DisplayInteractive -Description $item.Description
        if ($null -eq $selectedName) {
            Write-Host "  [SKIP] No display selected" -ForegroundColor Yellow
            $skipCount++
            Write-Host ""
            continue
        }
        $keyNamesToProcess = @($selectedName)
    }
    else {
        if ($target.MatchedKeyNames.Count -eq 0) {
            Write-Host "[$index] $($item.HardwareID)" -ForegroundColor Red
            Write-Host "  [ERROR] No display found matching '$($item.HardwareID)'" -ForegroundColor Red
            $failCount++
            Write-Host ""
            continue
        }
        $keyNamesToProcess = $target.MatchedKeyNames
    }

    # Resolve scale percent
    $scalePercent = $null
    if ($target.InteractiveScale) {
        $scalePercent = Select-ScaleInteractive
        if ($null -eq $scalePercent) {
            Write-Host "  [SKIP] No scale selected" -ForegroundColor Yellow
            $skipCount++
            Write-Host ""
            continue
        }
    }
    else {
        $scalePercent = [int]$sp
    }

    $targetDpi = $script:DpiValueMap[$scalePercent]

    foreach ($keyName in $keyNamesToProcess) {
        $displayLabel = if ($target.InteractiveDisplay) { $keyName } else { $item.HardwareID }
        Write-Host "[$index] $displayLabel -> $scalePercent% (DpiValue: $targetDpi)" -ForegroundColor Cyan
        Write-Host "  Key: $keyName" -ForegroundColor Gray

        $result = Write-DpiValue -FullKeyName $keyName -DpiValue $targetDpi -HiveLoaded $hiveLoaded

        if ($result.HasError) {
            $failCount++
        }
        elseif (-not $result.HkcuChanged -and -not $result.HiveChanged) {
            $skipCount++
        }
        else {
            $successCount++
        }

        Write-Host ""
    }
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

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Execution Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items (Already configured)" -ForegroundColor $(if ($skipCount -gt 0) { "Gray" } else { "Green" })
Write-Host "  Failed:  $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
if ($successCount -gt 0) {
    Write-Host ""
    Write-Host "NOTE: Sign-out or restart may be required for changes to take effect." -ForegroundColor Yellow
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($failCount -eq 0 -and $successCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
