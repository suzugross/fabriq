# ========================================
# Browser Addon Configuration Script
# ========================================
# Force-install Chrome/Edge extensions via Group Policy (registry)
# by registering them in ExtensionInstallForcelist.
#
# [NOTES]
# - Requires administrator privileges
# - Browser restart or `gpupdate /force` may be required to take effect
# ========================================

Write-Host ""
Show-Separator
Write-Host "Browser Addon Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Local Helper Functions
# ========================================

function Resolve-ExtensionId {
    param([string]$InputValue)
    $InputValue = $InputValue.Trim()
    # Raw extension ID (32 chars, a-p only)
    if ($InputValue -match '^[a-p]{32}$') { return $InputValue }
    # Chrome Web Store URL: /detail/[optional-name-slug/]<id>
    if ($InputValue -match '/detail/(?:[^/]+/)?([a-p]{32})') { return $Matches[1] }
    # Fallback: any 32-char a-p sequence in the string
    if ($InputValue -match '([a-p]{32})') { return $Matches[1] }
    return $null
}

function Test-ExtensionInForcelist {
    param(
        [string]$RegPath,
        [string]$ExtensionId
    )
    try {
        if (-not (Test-Path $RegPath)) { return $false }
        $item = Get-Item $RegPath -ErrorAction SilentlyContinue
        if ($null -eq $item -or $null -eq $item.Property) { return $false }
        foreach ($name in $item.Property) {
            $val = (Get-ItemProperty -Path $RegPath -Name $name -ErrorAction SilentlyContinue).$name
            if ($val -like "$ExtensionId;*") { return $true }
        }
        return $false
    }
    catch { return $false }
}

function Get-NextForcelistIndex {
    param([string]$RegPath)
    if (-not (Test-Path $RegPath)) { return 1 }
    $props = (Get-Item $RegPath -ErrorAction SilentlyContinue).Property
    if ($null -eq $props -or $props.Count -eq 0) { return 1 }
    $nums = @($props | ForEach-Object { [int]$_ } | Sort-Object)
    return ($nums[-1] + 1)
}


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "browser_addon_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Browser", "ExtensionId", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load browser_addon_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Preprocessing (resolve IDs, validate Browser)
# ========================================
$resolvedItems = @()

foreach ($item in $enabledItems) {
    $entry = $item | Select-Object *
    $browser = $item.Browser.Trim()

    # Browser validation
    $regPath = switch ($browser) {
        'Chrome' { 'HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist' }
        'Edge'   { 'HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist' }
        default  { $null }
    }

    if ($null -eq $regPath) {
        $entry | Add-Member -NotePropertyName 'ResolvedId' -NotePropertyValue $null
        $entry | Add-Member -NotePropertyName 'RegPath'    -NotePropertyValue $null
        $entry | Add-Member -NotePropertyName 'IsInvalid'  -NotePropertyValue $true
        $entry | Add-Member -NotePropertyName 'ErrorReason' -NotePropertyValue "Unsupported browser: $browser (Chrome or Edge only)"
        $resolvedItems += $entry
        continue
    }

    # Resolve extension ID
    $resolvedId = Resolve-ExtensionId -InputValue $item.ExtensionId

    if ($null -eq $resolvedId) {
        $entry | Add-Member -NotePropertyName 'ResolvedId' -NotePropertyValue $null
        $entry | Add-Member -NotePropertyName 'RegPath'    -NotePropertyValue $regPath
        $entry | Add-Member -NotePropertyName 'IsInvalid'  -NotePropertyValue $true
        $entry | Add-Member -NotePropertyName 'ErrorReason' -NotePropertyValue "Cannot resolve extension ID from: $($item.ExtensionId)"
        $resolvedItems += $entry
        continue
    }

    $entry | Add-Member -NotePropertyName 'ResolvedId'  -NotePropertyValue $resolvedId
    $entry | Add-Member -NotePropertyName 'RegPath'     -NotePropertyValue $regPath
    $entry | Add-Member -NotePropertyName 'IsInvalid'   -NotePropertyValue $false
    $entry | Add-Member -NotePropertyName 'ErrorReason' -NotePropertyValue $null
    $resolvedItems += $entry
}


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Extensions" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $resolvedItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.ExtensionId }

    if ($item.IsInvalid) {
        Write-Host "  [ERROR] $displayName" -ForegroundColor Red
        Write-Host "    Browser:      $($item.Browser)" -ForegroundColor DarkGray
        Write-Host "    Reason:       $($item.ErrorReason)" -ForegroundColor Red
        Write-Host ""
        continue
    }

    $isRegistered = Test-ExtensionInForcelist -RegPath $item.RegPath -ExtensionId $item.ResolvedId
    $marker = if ($isRegistered) { "[Current]" } else { "[Change]" }
    $markerColor = if ($isRegistered) { "Gray" } else { "White" }

    Write-Host "  $marker $displayName" -ForegroundColor $markerColor
    Write-Host "    Browser:      $($item.Browser)" -ForegroundColor DarkGray
    Write-Host "    Extension ID: $($item.ResolvedId)" -ForegroundColor DarkGray
    Write-Host "    Registry:     $($item.RegPath)" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above browser extension policies?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Apply-settings loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $resolvedItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.ResolvedId }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Invalid entry
    if ($item.IsInvalid) {
        Show-Error "$($item.ErrorReason)"
        $failCount++
        Write-Host ""
        continue
    }

    # Idempotency check
    if (Test-ExtensionInForcelist -RegPath $item.RegPath -ExtensionId $item.ResolvedId) {
        Show-Skip "Already in forcelist"
        $skipCount++
        Write-Host ""
        continue
    }

    try {
        $value = "$($item.ResolvedId);https://clients2.google.com/service/update2/crx"

        # Create registry key if not exists
        if (-not (Test-Path $item.RegPath)) {
            Write-Host "  -> Creating registry key: $($item.RegPath)" -ForegroundColor Gray
            New-Item -Path $item.RegPath -Force | Out-Null
        }

        # Get next index
        $nextIndex = Get-NextForcelistIndex -RegPath $item.RegPath

        # Write the entry
        New-ItemProperty -Path $item.RegPath -Name $nextIndex -Value $value `
            -PropertyType String -Force -ErrorAction Stop | Out-Null

        Show-Success "Registered as entry #$nextIndex ($($item.Browser))"
        $successCount++
    }
    catch {
        Show-Error "Failed: $_"
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

foreach ($item in $resolvedItems) {
    if ($item.IsInvalid) { continue }

    $displayName = if ($item.Description) { $item.Description } else { $item.ResolvedId }
    $isRegistered = Test-ExtensionInForcelist -RegPath $item.RegPath -ExtensionId $item.ResolvedId

    if ($isRegistered) {
        Write-Host "  [VERIFIED] $displayName" -ForegroundColor Green
        $verifyPass++
    } else {
        Write-Host "  [VERIFY FAILED] $displayName" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
# Invalid rows are excluded from the verify loop, so with zero verifiable
# rows $verifyFail stays 0 and a bare ($verifyFail -eq 0) would report
# Verified=true for a run where every row failed validation. Zero
# verified checks means "nothing was verified" -> $null (not implemented),
# never $true.
$verified = if (($verifyPass + $verifyFail) -eq 0) { $null } else { $verifyFail -eq 0 }

# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Browser Addon Configuration Results" -Verified $verified)
