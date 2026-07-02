# ========================================
# Local Group Member Configuration Script
# ========================================
# Adds domain groups/users or local users to local groups
# based on CSV configuration (group_list.csv).
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "Local Group Member Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper Functions
# ========================================

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

function Get-LocalGroupMemberNames {
    param([string]$GroupName)
    # Returns member names as 'AUTHORITY\leaf' (or bare 'leaf') strings.
    # Get-LocalGroupMember throws on groups containing orphaned SIDs
    # (PS 5.1 bug, common on kitting benches after user/profile removal);
    # fall back to ADSI enumeration, which tolerates them.
    try {
        return @(Get-LocalGroupMember -Group $GroupName -ErrorAction Stop | ForEach-Object { [string]$_.Name })
    }
    catch {
        $grp = [ADSI]("WinNT://./" + $GroupName + ",group")
        return @($grp.psbase.Invoke('Members') | ForEach-Object {
            $adsPath = ([ADSI]$_).psbase.Path -replace '^WinNT://', ''
            $parts = $adsPath -split '/'
            if ($parts.Count -ge 2) { "$($parts[-2])\$($parts[-1])" } else { [string]$parts[-1] }
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
        $names = Get-LocalGroupMemberNames -GroupName $GroupName
        if (-not $names) { return $false }

        # Authorities that satisfy this row. Local principals must belong
        # to this computer; domain principals to the CSV Domain - accepted
        # as the FQDN as-is or its first label (NetBIOS approximation,
        # since Get-LocalGroupMember reports NetBIOS while the CSV carries
        # FQDN). A same-name principal under any OTHER authority must NOT
        # satisfy the check: the old leaf-only wildcard match did, and
        # silently skipped the real add (cross-authority false-PASS).
        $accepted = if ($MemberType -eq 'LocalUser' -or [string]::IsNullOrWhiteSpace($Domain)) {
            @($env:COMPUTERNAME, '')
        } else {
            @($Domain, $Domain.Split('.')[0])
        }

        foreach ($n in $names) {
            $leaf = $n.Split('\')[-1]
            if ($leaf -ine $MemberName) { continue }
            $authority = if ($n.Contains('\')) { $n.Split('\')[0] } else { '' }
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

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "group_list.csv"

$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) { return (New-ModuleResult -Status "Error" -Message "Failed to load group_list.csv") }
if ($items.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }
Write-Host ""

# ========================================
# List Settings with Idempotency Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Target Group Member Settings" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$index = 0
foreach ($item in $items) {
    $index++
    $memberDisplay = Build-MemberName -Item $item

    # Check if group exists
    $groupExists = Test-LocalGroupExists -GroupName $item.LocalGroup
    if (-not $groupExists) {
        $marker = "[ERROR]"
        $markerColor = "Red"
    }
    else {
        # Check if already a member (use resolved name/authority for CurrentUser)
        $checkName = if ($item.MemberType -eq 'CurrentUser') { (Resolve-CurrentUser).Split('\')[-1] } else { $item.MemberName }
        $checkDomain = if ($item.MemberType -eq 'CurrentUser') {
            $full = Resolve-CurrentUser
            if ($full.Contains('\')) { $full.Split('\')[0] } else { '' }
        } else { [string]$item.Domain }
        $exists = Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $checkName -MemberType $item.MemberType -Domain $checkDomain
        if ($exists) {
            $marker = "[Current]"
            $markerColor = "Gray"
        }
        else {
            $marker = "[Change]"
            $markerColor = "White"
        }
    }

    Write-Host "[$index] $($item.LocalGroup) <- $memberDisplay  $marker" -ForegroundColor $markerColor
    Write-Host "    Type: $($item.MemberType) | $($item.Description)" -ForegroundColor Gray
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above group member settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Apply Settings
# ========================================
Show-Info "Applying group member settings..."
Write-Host ""

$successCount = 0
$skipCount = 0
$failCount = 0
$verifyPass = 0
$verifyFail = 0

$index = 0
foreach ($item in $items) {
    $index++
    $memberDisplay = Build-MemberName -Item $item

    Write-Host "[$index/$($items.Count)] $($item.LocalGroup) <- $memberDisplay" -ForegroundColor Cyan

    # Check if local group exists
    if (-not (Test-LocalGroupExists -GroupName $item.LocalGroup)) {
        Show-Error "Local group '$($item.LocalGroup)' not found"
        $failCount++
        $verifyFail++
        Write-Host ""
        continue
    }

    # Idempotency check (use resolved name/authority for CurrentUser)
    $checkName = if ($item.MemberType -eq 'CurrentUser') { (Resolve-CurrentUser).Split('\')[-1] } else { $item.MemberName }
    $checkDomain = if ($item.MemberType -eq 'CurrentUser') {
        $full = Resolve-CurrentUser
        if ($full.Contains('\')) { $full.Split('\')[0] } else { '' }
    } else { [string]$item.Domain }
    if (Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $checkName -MemberType $item.MemberType -Domain $checkDomain) {
        Show-Skip "Already a member"
        $skipCount++
        $verifyPass++
        Write-Host ""
        continue
    }

    # Add member
    try {
        Add-LocalGroupMember -Group $item.LocalGroup -Member $memberDisplay -ErrorAction Stop
        Show-Success "Member added"
        $successCount++
        # Step 5.5: Post-Apply Verification - re-read membership via the same helper.
        if (Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $checkName -MemberType $item.MemberType -Domain $checkDomain) {
            $verifyPass++
        }
        else {
            Show-Warning "Verify: '$memberDisplay' not found in '$($item.LocalGroup)' after add"
            $verifyFail++
        }
    }
    catch {
        if ($_.FullyQualifiedErrorId -match 'MemberExists') {
            # The OS says the principal is already a member even though the
            # matcher could not see it (custom NetBIOS name, unreadable
            # enumeration). Trust the authoritative signal: idempotent skip.
            Show-Skip "Already a member (reported by OS on add)"
            $skipCount++
            $verifyPass++
        }
        else {
            Show-Error "$($_.Exception.Message)"
            $failCount++
            $verifyFail++
        }
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
$verified = if (($verifyPass + $verifyFail) -gt 0) { ($verifyFail -eq 0) } else { $null }
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results" -Verified $verified)
