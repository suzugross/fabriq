# ========================================
# Windows Optional Feature Configuration
# ========================================
# [PURPOSE]
# Enable or disable Windows Optional Features online via DISM
# (Enable-/Disable-WindowsOptionalFeature) based on a CSV manifest.
#
# [NOTES]
# - Requires administrator privileges.
# - Online (running OS) only; offline WIM scenarios are out of scope.
# - For NetFx3 (DisabledWithPayloadRemoved by default on Win10/11),
#   set Source to the matching Windows ISO sources/sxs path, or leave
#   Source empty to fall back to Windows Update.
# - Restart is never triggered by this module; profile-level __RESTART__
#   is responsible for reboot orchestration.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Windows Optional Feature Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "windows_feature_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Action", "FeatureName")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load windows_feature_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Prerequisite check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required for Windows feature operations"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# ========================================
# Local Helpers
# ========================================
function Resolve-FeatureSourcePath {
    # Resolves a CSV Source value to an absolute filesystem path.
    # Drive-letter (C:\) and UNC (\\server\share) inputs are returned as-is.
    # Anything else is treated as module-relative and joined to $ModuleRoot.
    # Leading / or \ is stripped before joining (so "/payload/dotnetfx35"
    # and "payload\dotnetfx35" both resolve under the module folder).
    param(
        [string]$Path,
        [string]$ModuleRoot
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ($Path -match '^[A-Za-z]:[\\/]' -or $Path.StartsWith('\\')) {
        return $Path
    }
    $rel = $Path -replace '^[\\/]+', ''
    return (Join-Path $ModuleRoot $rel)
}

function Test-FeatureSourcePath {
    # Validates that a resolved Source path points to a directory containing
    # at least one *.cab file. Used by the Step 3 dry-run preview to catch
    # missing or empty payload folders before invoking DISM.
    param([string]$ResolvedPath)
    if ([string]::IsNullOrWhiteSpace($ResolvedPath)) {
        return [PSCustomObject]@{ Valid = $false; Reason = 'empty'; CabCount = 0 }
    }
    if (-not (Test-Path -LiteralPath $ResolvedPath -PathType Container)) {
        return [PSCustomObject]@{ Valid = $false; Reason = 'not-found'; CabCount = 0 }
    }
    $cabs = @(Get-ChildItem -LiteralPath $ResolvedPath -Filter '*.cab' -File -ErrorAction SilentlyContinue)
    if ($cabs.Count -eq 0) {
        return [PSCustomObject]@{ Valid = $false; Reason = 'no-cab'; CabCount = 0 }
    }
    return [PSCustomObject]@{ Valid = $true; Reason = 'ok'; CabCount = $cabs.Count }
}

function Get-FeatureStateSafely {
    # Returns the current State string, or $null when the feature name is
    # unknown to DISM on this OS.
    param([string]$FeatureName)
    try {
        $feat = Get-WindowsOptionalFeature -Online -FeatureName $FeatureName -ErrorAction Stop
        return $feat.State.ToString()
    }
    catch {
        return $null
    }
}

function Test-FeatureStateApplied {
    # True if $State already represents the requested $Action (used both for
    # idempotent skip in Step 5 and for Post-Apply Verification in Step 5.5).
    # Pending states count as applied; payload-removed counts as Disabled.
    param([string]$Action, [string]$State)
    if ($null -eq $State) { return $false }
    switch ($Action) {
        'Enable'  { return ($State -eq 'Enabled' -or $State -eq 'EnablePending') }
        'Disable' { return ($State -eq 'Disabled' -or $State -eq 'DisablePending' -or $State -eq 'DisabledWithPayloadRemoved') }
    }
    return $false
}

# ========================================
# Step 3: Dry-run summary
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Features" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$plans = @()
foreach ($item in $enabledItems) {
    $action = if ($item.Action) { $item.Action.Trim() } else { '' }

    if ($action -ne 'Enable' -and $action -ne 'Disable') {
        $plans += [PSCustomObject]@{
            Item = $item; Action = $action; ResolvedSource = $null
            SourceCheck = $null; State = $null
            Marker = '[INVALID]'; MarkerColor = 'Red'
        }
        continue
    }

    $state    = Get-FeatureStateSafely -FeatureName $item.FeatureName
    $resolved = Resolve-FeatureSourcePath -Path $item.Source -ModuleRoot $PSScriptRoot
    $srcCheck = if ($resolved) { Test-FeatureSourcePath -ResolvedPath $resolved } else { $null }

    $marker = '[APPLY]'
    $markerColor = 'Yellow'
    if ($null -eq $state) {
        $marker = '[NOT FOUND]'
        $markerColor = 'Red'
    }
    elseif (Test-FeatureStateApplied -Action $action -State $state) {
        $marker = '[Current]'
        $markerColor = 'Gray'
    }
    elseif ($action -eq 'Enable' -and $state -eq 'DisabledWithPayloadRemoved' -and $null -eq $resolved) {
        $marker = '[NEEDS SOURCE OR WU]'
        $markerColor = 'Magenta'
    }
    elseif ($null -ne $srcCheck -and -not $srcCheck.Valid) {
        $marker = '[BAD SOURCE]'
        $markerColor = 'Red'
    }

    $plans += [PSCustomObject]@{
        Item = $item; Action = $action; ResolvedSource = $resolved
        SourceCheck = $srcCheck; State = $state
        Marker = $marker; MarkerColor = $markerColor
    }
}

foreach ($p in $plans) {
    $item = $p.Item
    $displayName = if ($item.Description) { $item.Description } else { $item.FeatureName }
    Write-Host "  $($p.Marker) $displayName" -ForegroundColor $p.MarkerColor
    Write-Host "    Action  : $($p.Action)" -ForegroundColor DarkGray
    Write-Host "    Feature : $($item.FeatureName)" -ForegroundColor DarkGray
    $stateText = if ($p.State) { $p.State } else { 'unknown (feature name not recognized)' }
    Write-Host "    Current : $stateText" -ForegroundColor DarkGray
    if ($p.ResolvedSource) {
        $laRaw = if ($null -ne $item.PSObject.Properties['LimitAccess']) { $item.LimitAccess } else { '' }
        $laText = if ([string]::IsNullOrWhiteSpace($laRaw)) { '1 (default for closed networks)' } else { $laRaw }
        Write-Host "    Source  : $($p.ResolvedSource)" -ForegroundColor DarkGray
        Write-Host "    LimitAccess : $laText" -ForegroundColor DarkGray
        if ($p.SourceCheck -and -not $p.SourceCheck.Valid) {
            Write-Host "    [WARN] Source check failed: $($p.SourceCheck.Reason)" -ForegroundColor Red
        }
        elseif ($p.SourceCheck -and $p.SourceCheck.Valid) {
            Write-Host "    Source check : OK ($($p.SourceCheck.CabCount) cab files found)" -ForegroundColor DarkGray
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($item.Source)) {
        Write-Host "    Source  : $($item.Source) (resolution returned empty)" -ForegroundColor Red
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above feature changes?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Apply loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($p in $plans) {
    $item = $p.Item
    $displayName = if ($item.Description) { $item.Description } else { $item.FeatureName }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    if ($p.Action -ne 'Enable' -and $p.Action -ne 'Disable') {
        Show-Error "Invalid Action '$($item.Action)' (must be Enable or Disable)"
        $failCount++
        Write-Host ""
        continue
    }

    if ($null -eq $p.State) {
        Show-Error "Unknown feature: $($item.FeatureName)"
        $failCount++
        Write-Host ""
        continue
    }

    if (Test-FeatureStateApplied -Action $p.Action -State $p.State) {
        Show-Skip "Already in target state: $($p.State)"
        $skipCount++
        Write-Host ""
        continue
    }

    if ($p.Action -eq 'Enable' -and $p.ResolvedSource -and -not $p.SourceCheck.Valid) {
        Show-Error "Source path unusable ($($p.SourceCheck.Reason)): $($p.ResolvedSource)"
        $failCount++
        Write-Host ""
        continue
    }

    if ($p.Action -eq 'Enable' -and $p.State -eq 'DisabledWithPayloadRemoved' -and $null -eq $p.ResolvedSource) {
        Show-Warning "Feature payload is removed; will attempt Windows Update fallback (may fail on closed networks)"
    }

    # LimitAccess: default to $true when Source is set (closed-network safe);
    # ignored when Source is empty (DISM has nothing to limit).
    $useLimitAccess = $false
    if ($p.ResolvedSource) {
        $laRaw = if ($null -ne $item.PSObject.Properties['LimitAccess']) { $item.LimitAccess } else { '' }
        if ([string]::IsNullOrWhiteSpace($laRaw)) {
            $useLimitAccess = $true
        }
        else {
            $useLimitAccess = ($laRaw -eq '1')
        }
    }

    $useAll = ($item.IncludeAllSubFeatures -eq '1')

    try {
        if ($p.Action -eq 'Enable') {
            $params = @{
                Online      = $true
                FeatureName = $item.FeatureName
                NoRestart   = $true
                ErrorAction = 'Stop'
            }
            if ($useAll)              { $params.All         = $true }
            if ($p.ResolvedSource)    { $params.Source      = $p.ResolvedSource }
            if ($useLimitAccess)      { $params.LimitAccess = $true }
            $result = Enable-WindowsOptionalFeature @params
        }
        else {
            $result = Disable-WindowsOptionalFeature -Online `
                -FeatureName $item.FeatureName -NoRestart -ErrorAction Stop
        }

        Show-Success "Applied: $($p.Action) $($item.FeatureName)"
        if ($result -and $result.RestartNeeded) {
            Show-Info "Restart required for full effect (orchestrate via profile __RESTART__)"
        }
        $successCount++
    }
    catch {
        Show-Error "Failed: $($item.FeatureName) : $_"
        # Recognize source-not-found errors and provide actionable guidance.
        # The Japanese error message is matched too because OS UI language
        # affects DISM exception text.
        $msg = $_.Exception.Message
        if ($msg -match 'source files could not be found' `
            -or $msg -match 'ソース ファイルを見つけることができません' `
            -or $msg -match '0x800F081F' -or $msg -match '0x800F0954') {
            Show-Warning "Hint: For payload-removed features (e.g. NetFx3), point Source to a sources/sxs folder from a Windows ISO matching the running OS build, or clear Source to use Windows Update."
        }
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying applied feature states..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($p in $plans) {
    $item = $p.Item
    if ($p.Action -ne 'Enable' -and $p.Action -ne 'Disable') { continue }

    $displayName = if ($item.Description) { $item.Description } else { $item.FeatureName }
    $current = Get-FeatureStateSafely -FeatureName $item.FeatureName

    if ($null -eq $current) {
        Write-Host "  [VERIFY FAILED] $displayName (feature not found)" -ForegroundColor Red
        $verifyFail++
        continue
    }

    if (Test-FeatureStateApplied -Action $p.Action -State $current) {
        Write-Host "  [VERIFIED] $displayName (State=$current)" -ForegroundColor Green
        $verifyPass++
    }
    else {
        Write-Host "  [VERIFY FAILED] $displayName (State=$current; expected $($p.Action) target)" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0 -and ($verifyPass + $verifyFail) -gt 0)

# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Windows Feature Configuration Results" -Verified $verified)
