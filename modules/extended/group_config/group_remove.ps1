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

function Resolve-PrincipalSid {
    param(
        [string]$MemberName,
        [string]$MemberType,
        [string]$Domain
    )
    # Resolves the expected principal to a SID string, or $null when the
    # account cannot be resolved (off-domain bench, unreachable DC, typo).
    # SID identity is name-form independent, which is the whole point:
    # Get-LocalGroupMember reports the domain in NetBIOS form, the CSV
    # carries the FQDN, and in migrated domains the NetBIOS name is NOT
    # the first DNS label - name-based matching cannot bridge that.
    try {
        $authority = if ($MemberType -eq 'LocalUser' -or [string]::IsNullOrWhiteSpace($Domain)) {
            $env:COMPUTERNAME
        } else {
            $Domain
        }
        $account = New-Object System.Security.Principal.NTAccount($authority, $MemberName)
        return $account.Translate([System.Security.Principal.SecurityIdentifier]).Value
    }
    catch {
        return $null
    }
}

function Get-LocalGroupMemberIdentities {
    param([string]$GroupName)
    # Returns members as objects: Name = 'AUTHORITY\leaf' (or bare 'leaf'),
    # Sid = SID string or $null when unreadable.
    # Get-LocalGroupMember throws on groups containing orphaned SIDs
    # (PS 5.1 bug, common on kitting benches after user/profile removal);
    # fall back to ADSI enumeration, which tolerates them.
    try {
        return @(Get-LocalGroupMember -Group $GroupName -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{
                Name = [string]$_.Name
                Sid  = if ($_.SID) { [string]$_.SID.Value } else { $null }
            }
        })
    }
    catch {
        $grp = [ADSI]("WinNT://./" + $GroupName + ",group")
        return @($grp.psbase.Invoke('Members') | ForEach-Object {
            $entry = [ADSI]$_
            $adsPath = $entry.psbase.Path -replace '^WinNT://', ''
            $parts = $adsPath -split '/'
            $name = if ($parts.Count -ge 2) { "$($parts[-2])\$($parts[-1])" } else { [string]$parts[-1] }
            $sid = $null
            try {
                $sidBytes = $entry.psbase.InvokeGet('objectSid')
                if ($sidBytes) {
                    $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)).Value
                }
            } catch { }
            [pscustomobject]@{ Name = $name; Sid = $sid }
        })
    }
}

function Test-LocalGroupMemberExists {
    param(
        [string]$GroupName,
        [string]$MemberName,
        [string]$MemberType,
        [string]$Domain
    )
    try {
        $members = Get-LocalGroupMemberIdentities -GroupName $GroupName
        if (-not $members) { return $false }

        # Layer 1: SID identity (authoritative when both sides resolve).
        # For removal this is load-bearing: with a custom-NetBIOS domain
        # the name layer cannot see the member, the pre-check reported
        # "Already not a member" and the removal silently never happened.
        $expectedSid = Resolve-PrincipalSid -MemberName $MemberName -MemberType $MemberType -Domain $Domain

        # Layer 2 (fallback): same authority-scoped name matching as
        # group_config.ps1 for members whose SID could not be read or
        # when the expected SID is unresolvable. A same-name principal
        # under another authority must not match - prevents targeting the
        # wrong principal and stops the absence verification from
        # false-FAILing on an unrelated same-name member that remains.
        $accepted = if ($MemberType -eq 'LocalUser' -or [string]::IsNullOrWhiteSpace($Domain)) {
            @($env:COMPUTERNAME, '')
        } else {
            @($Domain, $Domain.Split('.')[0])
        }

        foreach ($m in $members) {
            if ($expectedSid -and $m.Sid -and ($m.Sid -eq $expectedSid)) { return $true }

            $leaf = $m.Name.Split('\')[-1]
            if ($leaf -ine $MemberName) { continue }
            $authority = if ($m.Name.Contains('\')) { $m.Name.Split('\')[0] } else { '' }
            foreach ($a in $accepted) {
                if ($authority -ieq $a) { return $true }
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
        # Check if currently a member (use resolved name/authority for CurrentUser)
        $checkName = if ($item.MemberType -eq 'CurrentUser') { (Resolve-CurrentUser).Split('\')[-1] } else { $item.MemberName }
        $checkDomain = if ($item.MemberType -eq 'CurrentUser') {
            $full = Resolve-CurrentUser
            if ($full.Contains('\')) { $full.Split('\')[0] } else { '' }
        } else { [string]$item.Domain }
        $exists = Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $checkName -MemberType $item.MemberType -Domain $checkDomain
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

    # Idempotency check (use resolved name/authority for CurrentUser)
    $checkName = if ($item.MemberType -eq 'CurrentUser') { (Resolve-CurrentUser).Split('\')[-1] } else { $item.MemberName }
    $checkDomain = if ($item.MemberType -eq 'CurrentUser') {
        $full = Resolve-CurrentUser
        if ($full.Contains('\')) { $full.Split('\')[0] } else { '' }
    } else { [string]$item.Domain }
    if (-not (Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $checkName -MemberType $item.MemberType -Domain $checkDomain)) {
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
        if ($_.FullyQualifiedErrorId -match 'MemberNotFound|PrincipalNotFound') {
            # The OS says the principal is not a member even though the
            # matcher saw a candidate (authority edge). End state = absent,
            # which is this row's goal: idempotent skip, not a failure.
            Show-Skip "Already not a member (reported by OS on remove)"
            $skipCount++
        }
        else {
            Show-Error "$($_.Exception.Message)"
            $failCount++
        }
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")
