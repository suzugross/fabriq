# ========================================
# Display Resolution Configuration Script
# ========================================
# Modifies display resolution via registry
# (GraphicsDrivers\Configuration) based on
# CSV configuration (display_list.csv).
# Restart required for changes to take effect.
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "Display Resolution Configuration" -ForegroundColor Cyan
Show-Separator
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
        Show-Error "No display configuration keys found in registry"
        return $null
    }

    Write-Host ""
    Show-Separator
    Write-Host "Available Displays" -ForegroundColor Cyan
    Show-Separator
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
        Show-Info "Invalid input, skipping"
        return $null
    }

    if ($selNum -eq 0 -or $selNum -gt $displayList.Count) {
        Show-Info "Skipped"
        return $null
    }

    return $displayList[$selNum - 1].Key
}

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "display_list.csv"

$allItems = Import-CsvSafe -Path $csvPath -Description "display_list.csv"
if ($null -eq $allItems -or $allItems.Count -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load display_list.csv")
}

# Filter enabled entries
$items = @($allItems | Where-Object { $_.Enabled -eq "1" })

if ($items.Count -eq 0) {
    Show-Skip "No enabled entries in display_list.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

Show-Info "Loaded $($items.Count) enabled entries (total: $($allItems.Count))"
Write-Host ""

# ========================================
# Validate CSV Data
# ========================================
$validItems = @()
foreach ($item in $items) {
    $w = 0; $h = 0
    if (-not [int]::TryParse($item.Width, [ref]$w) -or -not [int]::TryParse($item.Height, [ref]$h)) {
        Show-Warning "Invalid Width/Height for '$($item.Description)': Width=$($item.Width), Height=$($item.Height) — skipping"
        continue
    }
    if ($w -le 0 -or $h -le 0) {
        Show-Warning "Width/Height must be positive for '$($item.Description)' — skipping"
        continue
    }
    $validItems += $item
}

if ($validItems.Count -eq 0) {
    Show-Error "No valid entries after validation"
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
        Write-Host "[$index] $($item.HardwareID) -> $targetW x $targetH  [ERROR]"
        Write-Host "    No display found matching '$($item.HardwareID)'"
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
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Apply Settings
# ========================================
Show-Info "Applying Display Resolution Settings..."
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
            Show-Skip "No display selected"
            $skipCount++
            Write-Host ""
            continue
        }
        $keysToProcess = @($selectedKey)
    }
    else {
        if ($target.MatchedKeys.Count -eq 0) {
            Write-Host "[$index] $($item.HardwareID) -> $targetW x $targetH"
            Show-Error "No display found matching '$($item.HardwareID)'"
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
            Show-Error "No subkeys found under configuration key"
            $failCount++
            Write-Host ""
            continue
        }

        $subKey = $subKeys | Select-Object -First 1
        $subKeyPath = $subKey.PSPath

        # Check current values
        $props = Get-ItemProperty $subKeyPath -ErrorAction SilentlyContinue
        if ($null -eq $props) {
            Show-Error "Cannot read subkey properties"
            $failCount++
            Write-Host ""
            continue
        }

        $currentCx = $props.'PrimSurfSize.cx'
        $currentCy = $props.'PrimSurfSize.cy'

        if ($null -eq $currentCx -or $null -eq $currentCy) {
            Show-Error "PrimSurfSize.cx/cy not found in registry"
            $failCount++
            Write-Host ""
            continue
        }

        # Idempotency check
        if ([int]$currentCx -eq $targetW -and [int]$currentCy -eq $targetH) {
            Show-Skip "Already $targetW x $targetH"
            $skipCount++
            Write-Host ""
            continue
        }

        # Apply changes
        try {
            Set-ItemProperty -Path $subKeyPath -Name 'PrimSurfSize.cx' -Value $targetW -Type DWord -ErrorAction Stop
            Set-ItemProperty -Path $subKeyPath -Name 'PrimSurfSize.cy' -Value $targetH -Type DWord -ErrorAction Stop
            Show-Success "Changed from ${currentCx}x${currentCy} to ${targetW}x${targetH}"
            $successCount++
        }
        catch {
            Show-Error "$($_.Exception.Message)"
            $failCount++
        }

        Write-Host ""
    }
}

# ========================================
# Result Summary
# ========================================
Show-Separator
Write-Host "Execution Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items (Already configured)" -ForegroundColor $(if ($skipCount -gt 0) { "Gray" } else { "Green" })
Write-Host "  Failed:  $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Show-Separator
if ($successCount -gt 0) {
    Write-Host ""
    Show-Warning "Restart required for changes to take effect."
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($failCount -eq 0 -and $successCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    else { "Error" }
$restartNote = if ($successCount -gt 0) { " (Restart required)" } else { "" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount$restartNote")
