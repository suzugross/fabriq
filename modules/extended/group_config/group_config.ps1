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

function Build-MemberName {
    param([PSCustomObject]$Item)
    switch ($Item.MemberType) {
        'DomainGroup' { return "$($Item.Domain)\$($Item.MemberName)" }
        'DomainUser'  { return "$($Item.Domain)\$($Item.MemberName)" }
        'LocalUser'   { return $Item.MemberName }
        default       { return "$($Item.Domain)\$($Item.MemberName)" }
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
            if ($MemberType -eq 'LocalUser') {
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
        # Check if already a member
        $exists = Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $item.MemberName -MemberType $item.MemberType
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

$index = 0
foreach ($item in $items) {
    $index++
    $memberDisplay = Build-MemberName -Item $item

    Write-Host "[$index/$($items.Count)] $($item.LocalGroup) <- $memberDisplay" -ForegroundColor Cyan

    # Check if local group exists
    if (-not (Test-LocalGroupExists -GroupName $item.LocalGroup)) {
        Show-Error "Local group '$($item.LocalGroup)' not found"
        $failCount++
        Write-Host ""
        continue
    }

    # Idempotency check
    if (Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $item.MemberName -MemberType $item.MemberType) {
        Show-Skip "Already a member"
        $skipCount++
        Write-Host ""
        continue
    }

    # Add member
    try {
        Add-LocalGroupMember -Group $item.LocalGroup -Member $memberDisplay -ErrorAction Stop
        Show-Success "Member added"
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
