# ========================================
# Network Profile Configuration Script
# ========================================
# Dynamically resolves NetworkList profile GUID keys by ProfileName
# matching, and writes the Category value (0=Public, 1=Private, 2=Domain)
# to HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles\{GUID}.
#
# [NOTES]
# - Requires administrator privileges (HKLM write)
# - MatchMode=All applies to every existing profile (ProfileName ignored)
# - Rows are applied in CSV order; later rows override earlier ones
#   for profiles that match multiple rows
# - Post-Apply Verification checks each unique profile against its
#   final expected Category (after override resolution)
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "Network Profile Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Constants
# ========================================
$PROFILES_ROOT = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles'
$VALID_CATEGORIES = @{ '0' = 'Public'; '1' = 'Private'; '2' = 'Domain' }
$VALID_MATCH_MODES = @('Exact', 'Wildcard', 'All')


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "network_profile_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "ProfileName", "Category", "MatchMode")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load network_profile_list.csv")
}
if ($enabledItems.Count -eq 0) {
    Show-Info "No enabled entries in network_profile_list.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Normalize MatchMode (empty => Exact)
foreach ($item in $enabledItems) {
    if ([string]::IsNullOrWhiteSpace($item.MatchMode)) {
        $item.MatchMode = 'Exact'
    }
}


# ========================================
# Step 2: Prerequisites & Validation
# ========================================
if (-not (Test-Path $PROFILES_ROOT)) {
    Show-Error "Registry path not found: $PROFILES_ROOT"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "NetworkList\Profiles key not found")
}

$validationFailed = $false

foreach ($item in $enabledItems) {
    if (-not $VALID_CATEGORIES.ContainsKey([string]$item.Category)) {
        Show-Error "Invalid Category '$($item.Category)' (expected 0/1/2): $($item.Description)"
        $validationFailed = $true
    }
    if ($VALID_MATCH_MODES -notcontains $item.MatchMode) {
        Show-Error "Invalid MatchMode '$($item.MatchMode)' (expected Exact/Wildcard/All): $($item.Description)"
        $validationFailed = $true
    }
}

if ($validationFailed) {
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "CSV validation failed")
}

# Warn on MatchMode=All with non-empty ProfileName
$allRows = @($enabledItems | Where-Object { $_.MatchMode -eq 'All' })
foreach ($row in $allRows) {
    if (-not [string]::IsNullOrWhiteSpace($row.ProfileName)) {
        Show-Warning "MatchMode=All row has non-empty ProfileName '$($row.ProfileName)' - ProfileName will be ignored"
    }
}
if ($allRows.Count -gt 1) {
    Show-Warning "Multiple MatchMode=All rows detected ($($allRows.Count)) - later row wins for overlapping profiles"
}

Write-Host ""


# ========================================
# Helper: Test-CategoryValueMatch
# ========================================
# Idempotency check. Returns $true when the profile's Category equals
# the expected value, $false otherwise (including missing value).
function Test-CategoryValueMatch {
    param(
        [string]$ProfilePath,
        [int]$ExpectedCategory
    )
    try {
        $prop = Get-ItemProperty -Path $ProfilePath -Name Category -ErrorAction SilentlyContinue
        if ($null -eq $prop) { return $false }
        return ([int]$prop.Category -eq $ExpectedCategory)
    }
    catch {
        return $false
    }
}


# ========================================
# Helper: Resolve-NetworkProfileKeys
# ========================================
# Enumerates GUID subkeys under NetworkList\Profiles and returns those
# whose ProfileName matches the given pattern under the given MatchMode.
function Resolve-NetworkProfileKeys {
    param(
        [string]$ProfileName,
        [string]$MatchMode
    )
    $found = @()
    $subkeys = Get-ChildItem -Path $PROFILES_ROOT -ErrorAction SilentlyContinue
    foreach ($sk in $subkeys) {
        $props = Get-ItemProperty -Path $sk.PSPath -ErrorAction SilentlyContinue
        if ($null -eq $props) { continue }
        $name = $props.ProfileName
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $isMatch = switch ($MatchMode) {
            'All'      { $true }
            'Wildcard' { $name -like $ProfileName }
            default    { $name -eq $ProfileName }
        }

        if ($isMatch) {
            $currentCat = if ($null -ne $props.Category) { [int]$props.Category } else { -1 }
            $found += [pscustomobject]@{
                Path            = $sk.PSPath
                ProfileName     = $name
                Guid            = $sk.PSChildName
                CurrentCategory = $currentCat
            }
        }
    }
    return ,$found
}


# ========================================
# Step 3: Dry-run display
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following network category changes will be applied" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$resolvedPlan = @()

foreach ($item in $enabledItems) {
    $target = [int]$item.Category
    $targetLabel = $VALID_CATEGORIES["$target"]
    $displayPattern = if ([string]::IsNullOrWhiteSpace($item.ProfileName)) { "(all)" } else { $item.ProfileName }
    $displayName = if ($item.Description) { $item.Description } else { "$displayPattern [$($item.MatchMode)]" }

    Write-Host "[$($item.MatchMode)] $displayName -> Category=$target ($targetLabel)" -ForegroundColor Cyan

    $profileMatches = Resolve-NetworkProfileKeys -ProfileName $item.ProfileName -MatchMode $item.MatchMode

    if ($profileMatches.Count -eq 0) {
        Write-Host "  [NOT FOUND] No matching profile in registry" -ForegroundColor DarkYellow
        Write-Host ""
        $resolvedPlan += [pscustomobject]@{
            Row     = $item
            Matches = @()
        }
        continue
    }

    foreach ($m in $profileMatches) {
        $currentLabel = if ($VALID_CATEGORIES.ContainsKey("$($m.CurrentCategory)")) {
            $VALID_CATEGORIES["$($m.CurrentCategory)"]
        } else {
            "Unknown"
        }
        if ($m.CurrentCategory -eq $target) {
            Write-Host "  [CURRENT] $($m.ProfileName) ($($m.Guid)) already Category=$target ($currentLabel)" -ForegroundColor Gray
        } else {
            Write-Host "  [APPLY]   $($m.ProfileName) ($($m.Guid)) : $($m.CurrentCategory) ($currentLabel) -> $target ($targetLabel)" -ForegroundColor White
        }
    }
    Write-Host ""

    $resolvedPlan += [pscustomobject]@{
        Row     = $item
        Matches = $profileMatches
    }
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above network category changes?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""
Show-Info "Applying network category settings..."
Write-Host ""


# ========================================
# Step 5: Apply
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

# Final expected Category per profile path (later rows override earlier ones)
$finalExpected = [ordered]@{}

foreach ($plan in $resolvedPlan) {
    $item = $plan.Row
    $target = [int]$item.Category
    $displayPattern = if ([string]::IsNullOrWhiteSpace($item.ProfileName)) { "(all)" } else { $item.ProfileName }
    $displayName = if ($item.Description) { $item.Description } else { "$displayPattern [$($item.MatchMode)]" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "[$($item.MatchMode)] $displayName -> Category=$target" -ForegroundColor Yellow
    Write-Host "----------------------------------------" -ForegroundColor White

    if ($plan.Matches.Count -eq 0) {
        Show-Skip "No matching profile in registry"
        $skipCount++
        Write-Host ""
        continue
    }

    foreach ($m in $plan.Matches) {
        Write-Host "  Target: $($m.ProfileName) ($($m.Guid))" -ForegroundColor DarkGray

        try {
            if (Test-CategoryValueMatch -ProfilePath $m.Path -ExpectedCategory $target) {
                Show-Skip "  Already Category=$target"
                $skipCount++
                $finalExpected[$m.Path] = @{ ProfileName = $m.ProfileName; Target = $target }
                continue
            }

            Set-ItemProperty -Path $m.Path -Name Category -Value $target -Type DWord -Force -ErrorAction Stop
            Show-Success "  Category set to $target"
            $successCount++
            $finalExpected[$m.Path] = @{ ProfileName = $m.ProfileName; Target = $target }
        }
        catch {
            Show-Error "  $_"
            $failCount++
        }
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

foreach ($path in $finalExpected.Keys) {
    $info = $finalExpected[$path]
    if (Test-CategoryValueMatch -ProfilePath $path -ExpectedCategory $info.Target) {
        Write-Host "  [VERIFIED] $($info.ProfileName) -> Category=$($info.Target)" -ForegroundColor Green
        $verifyPass++
    } else {
        Write-Host "  [VERIFY FAILED] $($info.ProfileName) expected Category=$($info.Target)" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = if (($verifyPass + $verifyFail) -gt 0) { ($verifyFail -eq 0) } else { $null }


# ========================================
# Step 6: Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Network Profile Results" -Verified $verified)
