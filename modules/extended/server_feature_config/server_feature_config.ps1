# ========================================
# Server Feature Install Script (Windows Server)
# ========================================
# [PURPOSE]
# Install Windows Server roles, role services, and features online via
# ServerManager (Install-WindowsFeature) based on a CSV manifest.
#
# [NOTES]
# - Requires administrator privileges.
# - Windows Server only (Get-CimInstance Win32_OperatingSystem ProductType
#   in {2, 3}). Client SKUs are reported as Skipped by design so the same
#   profile can run safely on mixed fleets.
# - Online (running OS) only; offline VHD and ConfigurationFilePath
#   scenarios are intentionally out of scope.
# - Restart is never triggered by this module; profile-level __RESTART__
#   is responsible for reboot orchestration.
# - For features with payload removed, set Source to a sources/sxs path
#   on the matching Windows Server media or a UNC share, or leave Source
#   empty to fall back to Windows Update.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Server Feature Install" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "server_feature_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Name", "IncludeAllSubFeature", "IncludeManagementTools")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load server_feature_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Prerequisite checks
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required for server feature operations"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# Windows Server-only guard. ProductType: 1 = Workstation (client SKU),
# 2 = Domain Controller, 3 = Server. Returning Skipped (not Error) keeps
# mixed-fleet profiles from failing on client targets.
$productType = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).ProductType
if ($productType -eq 1) {
    Show-Skip "This module is for Windows Server only (current OS is a client SKU)"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Client SKU detected; Server-only module")
}

try {
    Import-Module ServerManager -ErrorAction Stop
}
catch {
    Show-Error "Failed to load ServerManager module: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "ServerManager module unavailable")
}

# ========================================
# Local Helpers
# ========================================
function Resolve-FeatureSourcePath {
    # Resolves a CSV Source value to an absolute filesystem path.
    # Drive-letter (C:\) and UNC (\\server\share) inputs are returned as-is.
    # Anything else is treated as module-relative and joined to $ModuleRoot.
    # Leading / or \ is stripped before joining, mirroring the convention
    # used by the sister module windows_feature_config.
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

function Get-FeatureStateSafely {
    # Returns the InstallState string (Available / Installed / InstallPending /
    # Removed), or $null when the feature name is unknown to ServerManager
    # on this OS. Wrapped to absorb cmdlet warnings and missing-feature
    # error variants without throwing past the caller.
    param([string]$FeatureName)
    try {
        $results = @(Get-WindowsFeature -Name $FeatureName `
                        -ErrorAction Stop -WarningAction SilentlyContinue)
        if ($results.Count -eq 0) { return $null }
        $first = $results[0]
        if ($null -eq $first) { return $null }
        return $first.InstallState.ToString()
    }
    catch {
        return $null
    }
}

function Test-InstallStateApplied {
    # True when the current state already represents an installed feature.
    # InstallPending is treated as applied (accept-as-applied when restart
    # is the only remaining step). Same approach as hostname_config and
    # windows_feature_config Step 5.5.
    param([string]$State)
    if ($null -eq $State) { return $false }
    return ($State -eq 'Installed' -or $State -eq 'InstallPending')
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
    $state    = Get-FeatureStateSafely -FeatureName $item.Name
    $resolved = Resolve-FeatureSourcePath -Path $item.Source -ModuleRoot $PSScriptRoot

    $marker = '[INSTALL]'
    $markerColor = 'Yellow'
    if ($null -eq $state) {
        $marker = '[NOT FOUND]'
        $markerColor = 'Red'
    }
    elseif ($state -eq 'Installed') {
        $marker = '[ALREADY-INSTALLED]'
        $markerColor = 'Gray'
    }
    elseif ($state -eq 'InstallPending') {
        $marker = '[PENDING-RESTART]'
        $markerColor = 'DarkYellow'
    }
    elseif ($state -eq 'Removed' -and $null -eq $resolved) {
        $marker = '[NEEDS SOURCE OR WU]'
        $markerColor = 'Magenta'
    }

    $plans += [PSCustomObject]@{
        Item           = $item
        ResolvedSource = $resolved
        State          = $state
        Marker         = $marker
        MarkerColor    = $markerColor
    }
}

foreach ($p in $plans) {
    $item = $p.Item
    $displayName = if ($item.Description) { $item.Description } else { $item.Name }
    Write-Host "  $($p.Marker) $displayName" -ForegroundColor $p.MarkerColor
    Write-Host "    Name    : $($item.Name)" -ForegroundColor DarkGray
    $stateText = if ($p.State) { $p.State } else { 'unknown (feature name not recognized on this Server SKU)' }
    Write-Host "    Current : $stateText" -ForegroundColor DarkGray
    $allText  = if ($item.IncludeAllSubFeature   -eq '1') { 'yes' } else { 'no' }
    $mgmtText = if ($item.IncludeManagementTools -eq '1') { 'yes' } else { 'no' }
    Write-Host "    IncludeAllSubFeature   : $allText"  -ForegroundColor DarkGray
    Write-Host "    IncludeManagementTools : $mgmtText" -ForegroundColor DarkGray
    if ($p.ResolvedSource) {
        Write-Host "    Source  : $($p.ResolvedSource)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Install the listed Windows Server features?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Apply loop
# ========================================
$successCount       = 0
$skipCount          = 0
$failCount          = 0
$restartNeededCount = 0

foreach ($p in $plans) {
    $item = $p.Item
    $displayName = if ($item.Description) { $item.Description } else { $item.Name }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    if ($null -eq $p.State) {
        Show-Error "Unknown feature: $($item.Name)"
        $failCount++
        Write-Host ""
        continue
    }

    if (Test-InstallStateApplied -State $p.State) {
        Show-Skip "Already installed (State=$($p.State))"
        $skipCount++
        Write-Host ""
        continue
    }

    if ($p.State -eq 'Removed' -and $null -eq $p.ResolvedSource) {
        Show-Warning "Feature payload is removed; will attempt Windows Update fallback (may fail on closed networks)"
    }

    $params = @{
        Name        = $item.Name
        ErrorAction = 'Stop'
    }
    if ($item.IncludeAllSubFeature   -eq '1') { $params.IncludeAllSubFeature   = $true }
    if ($item.IncludeManagementTools -eq '1') { $params.IncludeManagementTools = $true }
    if ($p.ResolvedSource)                    { $params.Source                 = $p.ResolvedSource }
    # -Restart is intentionally not passed: reboot is orchestrated by
    # profile-level __RESTART__ so AutoPilot flow stays deterministic.

    try {
        $result = Install-WindowsFeature @params
        if ($result -and $result.Success) {
            Show-Success "Installed: $($item.Name)"
            if ($result.RestartNeeded -eq 'Yes') {
                Show-Info "Restart required for full effect (orchestrate via profile __RESTART__)"
                $restartNeededCount++
            }
            $successCount++
        }
        else {
            $exitCode = if ($result) { $result.ExitCode } else { 'n/a' }
            Show-Error "Install reported failure: $($item.Name) (ExitCode=$exitCode)"
            $failCount++
        }
    }
    catch {
        Show-Error "Failed: $($item.Name) : $_"
        # Recognize source-not-found errors and provide actionable guidance.
        # Japanese error text is matched too because OS UI language affects
        # cmdlet exception text.
        $msg = $_.Exception.Message
        if ($msg -match 'source files could not be found' `
            -or $msg -match 'ソース ファイルを見つけることができません' `
            -or $msg -match '0x800F081F' -or $msg -match '0x800F0954') {
            Show-Warning "Hint: For payload-removed features, point Source to a sources/sxs folder from a Windows Server ISO matching the running OS build, or clear Source to use Windows Update."
        }
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying installed feature states..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($p in $plans) {
    $item = $p.Item
    $displayName = if ($item.Description) { $item.Description } else { $item.Name }
    $current = Get-FeatureStateSafely -FeatureName $item.Name

    if ($null -eq $current) {
        Write-Host "  [VERIFY FAILED] $displayName (feature not found)" -ForegroundColor Red
        $verifyFail++
        continue
    }

    if (Test-InstallStateApplied -State $current) {
        Write-Host "  [VERIFIED] $displayName (State=$current)" -ForegroundColor Green
        $verifyPass++
    }
    else {
        Write-Host "  [VERIFY FAILED] $displayName (State=$current; expected Installed/InstallPending)" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0 -and ($verifyPass + $verifyFail) -gt 0)

# ========================================
# Step 6: Aggregate and return result
# ========================================
$suffix = ""
if ($restartNeededCount -gt 0) {
    Show-Warning "Restart required for $restartNeededCount feature(s). Use __RESTART__ marker in profile."
    Write-Host ""
    $suffix = "(restart required for $restartNeededCount item(s))"
}

return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Server Feature Install Results" -MessageSuffix $suffix -Verified $verified)
