# ========================================
# Local Group Member Removal Script
# ========================================
# Removes domain groups/users or local users from local groups
# based on the shared CSV configuration (group_list.csv).
# Logical inverse of group_config.ps1 (member addition).
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "Local Group Member Removal" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper Functions
# ========================================
# Duplicated from group_config.ps1. Fabriq modules are self-contained
# (no sibling dot-sourcing convention), so the shared helpers are mirrored.

function Resolve-CurrentUser {
    # Get actual logged-on user via WMI (handles UAC elevation correctly)
    # Win32_ComputerSystem.UserName returns "DOMAIN\username" for the interactive session
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $loggedOnUser = $cs.UserName
        if (-not [string]::IsNullOrWhiteSpace($loggedOnUser)) {
            return $loggedOnUser
        }
    }
    catch { }
    # Fallback to environment variables
    if ($env:USERDOMAIN -ne $env:COMPUTERNAME) {
        return "$env:USERDOMAIN\$env:USERNAME"
    }
    return $env:USERNAME
}

function Build-MemberName {
    param([PSCustomObject]$Item)
    switch ($Item.MemberType) {
        'DomainGroup'  { return "$($Item.Domain)\$($Item.MemberName)" }
        'DomainUser'   { return "$($Item.Domain)\$($Item.MemberName)" }
        'LocalUser'    { return $Item.MemberName }
        'CurrentUser'  { return (Resolve-CurrentUser) }
        default        { return "$($Item.Domain)\$($Item.MemberName)" }
    }
}

function Test-LocalGroupMemberExists {
    param(
        [string]$GroupName,
        [string]$MemberName,
        [string]$MemberType
    )
    try {
        $members = Get-LocalGroupMember -Group $GroupName -ErrorAction Stop
        if (-not $members) { return $false }

        foreach ($m in $members) {
            $isLocalType = ($MemberType -eq 'LocalUser')
            if ($isLocalType) {
                if ($m.Name -eq "$env:COMPUTERNAME\$MemberName" -or $m.Name -eq $MemberName) {
                    return $true
                }
            }
            else {
                if ($m.Name -like "*\$MemberName") {
                    return $true
                }
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

function Test-LocalGroupExists {
    param([string]$GroupName)
    try {
        $null = Get-LocalGroup -Name $GroupName -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Test-IsCurrentUserTarget {
    param([PSCustomObject]$Item)
    # True when the row targets the current logged-on user's own identity.
    # Scope: direct user identity only (CurrentUser / DomainUser / LocalUser).
    # DomainGroup is NOT resolved here (would require an AD membership query).
    $currentShort = (Resolve-CurrentUser).Split('\')[-1]
    switch ($Item.MemberType) {
        'CurrentUser' { return $true }
        'DomainUser'  { return (-not [string]::IsNullOrWhiteSpace($Item.MemberName)) -and ($Item.MemberName -eq $currentShort) }
        'LocalUser'   { return (-not [string]::IsNullOrWhiteSpace($Item.MemberName)) -and ($Item.MemberName -eq $currentShort) }
        default       { return $false }
    }
}

# ========================================
# Load CSV (shared with group_config.ps1)
# ========================================
$csvPath = Join-Path $PSScriptRoot "group_list.csv"

$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) { return (New-ModuleResult -Status "Error" -Message "Failed to load group_list.csv") }
if ($items.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }
Write-Host ""

# ========================================
# List Settings with Idempotency Check (removal context)
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Target Group Member Removals" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$index = 0
foreach ($item in $items) {
    $index++
    $memberDisplay = Build-MemberName -Item $item
    $exists = $false

    # Check if group exists
    $groupExists = Test-LocalGroupExists -GroupName $item.LocalGroup
    if (-not $groupExists) {
        $marker = "[No Group]"
        $markerColor = "Gray"
    }
    else {
        # Check if currently a member (use resolved name for CurrentUser)
        $checkName = if ($item.MemberType -eq 'CurrentUser') { (Resolve-CurrentUser).Split('\')[-1] } else { $item.MemberName }
        $exists = Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $checkName -MemberType $item.MemberType
        if ($exists) {
            $marker = "[Remove]"
            $markerColor = "Yellow"
        }
        else {
            $marker = "[Absent]"
            $markerColor = "Gray"
        }
    }

    Write-Host "[$index] $($item.LocalGroup) -X $memberDisplay  $marker" -ForegroundColor $markerColor
    Write-Host "    Type: $($item.MemberType) | $($item.Description)" -ForegroundColor Gray

    # Self-removal warning (warn-only; surfaced before confirmation, does not block)
    if ($groupExists -and $exists -and (Test-IsCurrentUserTarget -Item $item)) {
        if ($item.LocalGroup -eq 'Administrators') {
            Show-Warning "Row [$index] removes the CURRENT logged-on user from Administrators; the elevated session may lose privileges"
        }
        else {
            Show-Warning "Row [$index] removes the current logged-on user from '$($item.LocalGroup)'"
        }
    }

    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Remove the above group members?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Apply Settings (removal)
# ========================================
Show-Info "Removing group members..."
Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0

$index = 0
foreach ($item in $items) {
    $index++
    $memberDisplay = Build-MemberName -Item $item

    Write-Host "[$index/$($items.Count)] $($item.LocalGroup) -X $memberDisplay" -ForegroundColor Cyan

    # Skip when the local group does not exist (nothing to remove)
    if (-not (Test-LocalGroupExists -GroupName $item.LocalGroup)) {
        Show-Skip "Local group '$($item.LocalGroup)' not found; nothing to remove"
        $skipCount++
        Write-Host ""
        continue
    }

    # Idempotency check (use resolved name for CurrentUser)
    $checkName = if ($item.MemberType -eq 'CurrentUser') { (Resolve-CurrentUser).Split('\')[-1] } else { $item.MemberName }
    if (-not (Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $checkName -MemberType $item.MemberType)) {
        Show-Skip "Already not a member"
        $skipCount++
        Write-Host ""
        continue
    }

    # Self-removal warning (warn-only; does not block, even under AutoPilot)
    if (Test-IsCurrentUserTarget -Item $item) {
        if ($item.LocalGroup -eq 'Administrators') {
            Show-Warning "Removing the CURRENT logged-on user from Administrators; the elevated session may lose privileges"
        }
        else {
            Show-Warning "Removing the current logged-on user from '$($item.LocalGroup)'"
        }
    }

    # Remove member
    try {
        Remove-LocalGroupMember -Group $item.LocalGroup -Member $memberDisplay -ErrorAction Stop
        Show-Success "Member removed"
        $successCount++
    }
    catch {
        Show-Error "$($_.Exception.Message)"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")
