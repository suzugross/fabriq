# ========================================
# Local Group Member Configuration Script
# ========================================
# Adds domain groups/users or local users to local groups
# based on CSV configuration (group_list.csv).
# ========================================

# Check Administrator Privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Administrator privileges are required."
    Write-Warning "Please run PowerShell as Administrator and try again."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Local Group Member Configuration" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
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

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] group_list.csv not found: $csvPath" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "group_list.csv not found")
}

try {
    $allItems = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load group_list.csv: $_" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to load group_list.csv: $_")
}

if ($allItems.Count -eq 0) {
    Write-Host "[ERROR] group_list.csv contains no data" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "group_list.csv contains no data")
}

# Filter enabled entries
$items = @($allItems | Where-Object { $_.Enabled -eq "1" })

if ($items.Count -eq 0) {
    Write-Host "[INFO] No enabled entries in group_list.csv" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

Write-Host "[INFO] Loaded $($items.Count) enabled entries (total: $($allItems.Count))" -ForegroundColor Cyan
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
if (-not (Confirm-Execution -Message "Apply the above group member settings?")) {
    Write-Host ""
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Apply Settings
# ========================================
Write-Host "--- Applying Group Member Settings ---" -ForegroundColor Cyan
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
        Write-Host "  [ERROR] Local group '$($item.LocalGroup)' not found" -ForegroundColor Red
        $failCount++
        Write-Host ""
        continue
    }

    # Idempotency check
    if (Test-LocalGroupMemberExists -GroupName $item.LocalGroup -MemberName $item.MemberName -MemberType $item.MemberType) {
        Write-Host "  [SKIP] Already a member" -ForegroundColor Gray
        $skipCount++
        Write-Host ""
        continue
    }

    # Add member
    try {
        Add-LocalGroupMember -Group $item.LocalGroup -Member $memberDisplay -ErrorAction Stop
        Write-Host "  [SUCCESS] Member added" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $failCount++
    }

    Write-Host ""
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
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($failCount -eq 0 -and $successCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
