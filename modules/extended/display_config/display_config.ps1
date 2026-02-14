# ========================================
# Display Resolution Configuration Script
# ========================================
# Modifies display resolution via registry
# (GraphicsDrivers\Configuration) based on
# CSV configuration (display_list.csv).
# Restart required for changes to take effect.
# ========================================

# Check Administrator Privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Administrator privileges are required."
    Write-Warning "Please run PowerShell as Administrator and try again."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Display Resolution Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Constants
# ========================================
$script:ConfigBasePath = 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration'

# ========================================
# Helper Functions
# ========================================

function Find-DisplayConfigKeys {
    param([string]$HardwareID)

    $allKeys = Get-ChildItem $script:ConfigBasePath -ErrorAction SilentlyContinue
    if (-not $allKeys) { return @() }

    if ([string]::IsNullOrWhiteSpace($HardwareID)) {
        return @($allKeys)
    }

    return @($allKeys | Where-Object { $_.PSChildName -like "$HardwareID*" })
}

function Get-DisplayCurrentResolution {
    param([Microsoft.Win32.RegistryKey]$ConfigKey)

    $subKeys = Get-ChildItem $ConfigKey.PSPath -ErrorAction SilentlyContinue
    if (-not $subKeys) { return $null }

    $subKey = $subKeys | Select-Object -First 1
    $props = Get-ItemProperty $subKey.PSPath -ErrorAction SilentlyContinue
    if ($null -eq $props) { return $null }

    $cx = $props.'PrimSurfSize.cx'
    $cy = $props.'PrimSurfSize.cy'

    if ($null -eq $cx -or $null -eq $cy) { return $null }

    return [PSCustomObject]@{
        Width      = [int]$cx
        Height     = [int]$cy
        SubKeyPath = $subKey.PSPath
    }
}

function Select-DisplayInteractive {
    param(
        [int]$Width,
        [int]$Height,
        [string]$Description
    )

    $allKeys = Find-DisplayConfigKeys -HardwareID ""
    if ($allKeys.Count -eq 0) {
        Write-Host "[ERROR] No display configuration keys found in registry" -ForegroundColor Red
        return $null
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Available Displays" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    $displayList = @()
    $idx = 0
    foreach ($key in $allKeys) {
        $idx++
        $res = Get-DisplayCurrentResolution -ConfigKey $key
        $currentStr = if ($res) { "$($res.Width) x $($res.Height)" } else { "Unknown" }

        Write-Host "[$idx] $($key.PSChildName)" -ForegroundColor White
        Write-Host "    Current: $currentStr" -ForegroundColor Gray
        Write-Host ""

        $displayList += [PSCustomObject]@{
            Index     = $idx
            Key       = $key
            Current   = $res
        }
    }

    Write-Host "Target resolution: $Width x $Height ($Description)" -ForegroundColor Cyan
    Write-Host ""

    $selection = Read-Host "Select display number (or 0 to skip)"
    $selNum = 0
    if (-not [int]::TryParse($selection, [ref]$selNum)) {
        Write-Host "[INFO] Invalid input, skipping" -ForegroundColor Yellow
        return $null
    }

    if ($selNum -eq 0 -or $selNum -gt $displayList.Count) {
        Write-Host "[INFO] Skipped" -ForegroundColor Yellow
        return $null
    }

    return $displayList[$selNum - 1].Key
}

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "display_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] display_list.csv not found: $csvPath" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "display_list.csv not found")
}

try {
    $allItems = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load display_list.csv: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to load display_list.csv: $_")
}

if ($allItems.Count -eq 0) {
    Write-Host "[ERROR] display_list.csv contains no data" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "display_list.csv contains no data")
}

# Filter enabled entries
$items = @($allItems | Where-Object { $_.Enabled -eq "1" })

if ($items.Count -eq 0) {
    Write-Host "[INFO] No enabled entries in display_list.csv" -ForegroundColor Yellow
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
    $w = 0; $h = 0
    if (-not [int]::TryParse($item.Width, [ref]$w) -or -not [int]::TryParse($item.Height, [ref]$h)) {
        Write-Host "[WARN] Invalid Width/Height for '$($item.Description)': Width=$($item.Width), Height=$($item.Height) — skipping" -ForegroundColor Yellow
        continue
    }
    if ($w -le 0 -or $h -le 0) {
        Write-Host "[WARN] Width/Height must be positive for '$($item.Description)' — skipping" -ForegroundColor Yellow
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
# Build a list of (item, matchedKeys, currentResolution) for each entry

$targets = @()

foreach ($item in $validItems) {
    $hwId = $item.HardwareID.Trim()

    if ([string]::IsNullOrWhiteSpace($hwId)) {
        # Interactive selection mode — deferred to apply phase
        $targets += [PSCustomObject]@{
            Item          = $item
            MatchedKeys   = @()
            Interactive   = $true
        }
    }
    else {
        $matched = Find-DisplayConfigKeys -HardwareID $hwId
        $targets += [PSCustomObject]@{
            Item          = $item
            MatchedKeys   = $matched
            Interactive   = $false
        }
    }
}

# ========================================
# List Settings with Idempotency Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Target Display Settings" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$index = 0
foreach ($target in $targets) {
    $index++
    $item = $target.Item
    $targetW = [int]$item.Width
    $targetH = [int]$item.Height

    if ($target.Interactive) {
        Write-Host "[$index] (Interactive Selection) -> $targetW x $targetH" -ForegroundColor Yellow
        Write-Host "    Display will be selected during apply phase" -ForegroundColor Gray
        Write-Host "    $($item.Description)" -ForegroundColor Gray
        Write-Host ""
        continue
    }

    if ($target.MatchedKeys.Count -eq 0) {
        Write-Host "[$index] $($item.HardwareID) -> $targetW x $targetH  [ERROR]" -ForegroundColor Red
        Write-Host "    No display found matching '$($item.HardwareID)'" -ForegroundColor Red
        Write-Host "    $($item.Description)" -ForegroundColor Gray
        Write-Host ""
        continue
    }

    foreach ($key in $target.MatchedKeys) {
        $res = Get-DisplayCurrentResolution -ConfigKey $key
        $currentStr = if ($res) { "$($res.Width) x $($res.Height)" } else { "Unknown" }

        if ($res -and $res.Width -eq $targetW -and $res.Height -eq $targetH) {
            $marker = "[Current]"
            $markerColor = "Gray"
        }
        else {
            $marker = "[Change]"
            $markerColor = "White"
        }

        Write-Host "[$index] $($item.HardwareID) -> $targetW x $targetH  $marker" -ForegroundColor $markerColor
        Write-Host "    Matched: $($key.PSChildName)" -ForegroundColor Gray
        Write-Host "    Current: $currentStr | $($item.Description)" -ForegroundColor Gray
        Write-Host ""
    }
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Apply the above display resolution settings?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Apply Settings
# ========================================
Write-Host "--- Applying Display Resolution Settings ---" -ForegroundColor Cyan
Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0

$index = 0
foreach ($target in $targets) {
    $index++
    $item = $target.Item
    $targetW = [int]$item.Width
    $targetH = [int]$item.Height

    # Handle interactive selection
    $keysToProcess = @()
    if ($target.Interactive) {
        $selectedKey = Select-DisplayInteractive -Width $targetW -Height $targetH -Description $item.Description
        if ($null -eq $selectedKey) {
            Write-Host "  [SKIP] No display selected" -ForegroundColor Yellow
            $skipCount++
            Write-Host ""
            continue
        }
        $keysToProcess = @($selectedKey)
    }
    else {
        if ($target.MatchedKeys.Count -eq 0) {
            Write-Host "[$index] $($item.HardwareID) -> $targetW x $targetH" -ForegroundColor Red
            Write-Host "  [ERROR] No display found matching '$($item.HardwareID)'" -ForegroundColor Red
            $failCount++
            Write-Host ""
            continue
        }
        $keysToProcess = $target.MatchedKeys
    }

    foreach ($configKey in $keysToProcess) {
        $displayName = if ($target.Interactive) { $configKey.PSChildName } else { $item.HardwareID }
        Write-Host "[$index] $displayName -> $targetW x $targetH" -ForegroundColor Cyan
        Write-Host "  Key: $($configKey.PSChildName)" -ForegroundColor Gray

        # Find subkey (00, 01, etc.)
        $subKeys = Get-ChildItem $configKey.PSPath -ErrorAction SilentlyContinue
        if (-not $subKeys) {
            Write-Host "  [ERROR] No subkeys found under configuration key" -ForegroundColor Red
            $failCount++
            Write-Host ""
            continue
        }

        $subKey = $subKeys | Select-Object -First 1
        $subKeyPath = $subKey.PSPath

        # Check current values
        $props = Get-ItemProperty $subKeyPath -ErrorAction SilentlyContinue
        if ($null -eq $props) {
            Write-Host "  [ERROR] Cannot read subkey properties" -ForegroundColor Red
            $failCount++
            Write-Host ""
            continue
        }

        $currentCx = $props.'PrimSurfSize.cx'
        $currentCy = $props.'PrimSurfSize.cy'

        if ($null -eq $currentCx -or $null -eq $currentCy) {
            Write-Host "  [ERROR] PrimSurfSize.cx/cy not found in registry" -ForegroundColor Red
            $failCount++
            Write-Host ""
            continue
        }

        # Idempotency check
        if ([int]$currentCx -eq $targetW -and [int]$currentCy -eq $targetH) {
            Write-Host "  [SKIP] Already $targetW x $targetH" -ForegroundColor Gray
            $skipCount++
            Write-Host ""
            continue
        }

        # Apply changes
        try {
            Set-ItemProperty -Path $subKeyPath -Name 'PrimSurfSize.cx' -Value $targetW -Type DWord -ErrorAction Stop
            Set-ItemProperty -Path $subKeyPath -Name 'PrimSurfSize.cy' -Value $targetH -Type DWord -ErrorAction Stop
            Write-Host "  [SUCCESS] Changed from ${currentCx}x${currentCy} to ${targetW}x${targetH}" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
            $failCount++
        }

        Write-Host ""
    }
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
    Write-Host "NOTE: Restart required for changes to take effect." -ForegroundColor Yellow
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($failCount -eq 0 -and $successCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    else { "Error" }
$restartNote = if ($successCount -gt 0) { " (Restart required)" } else { "" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount$restartNote")
