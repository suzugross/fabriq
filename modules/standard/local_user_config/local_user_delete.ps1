# ========================================
# Local User Deletion Script
# ========================================

Show-Info "Executing local user deletion..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "local_user_list.csv"

$userList = Import-ModuleCsv -Path $csvPath
if ($null -eq $userList) {
    # PowerShell auto-unwraps @() returned for "Segment 0 match" into $null; treat as empty so host CSV merge can still supply users.
    if (-not (Test-Path $csvPath)) {
        return (New-ModuleResult -Status "Error" -Message "local_user_list.csv not found")
    }
    $userList = @()
}

# ========================================
# Load Per-PC User CSV (optional)
# ========================================
$hostCsvPath = Join-Path $PSScriptRoot "local_user_host_list.csv"

if ((Test-Path $hostCsvPath) -and -not [string]::IsNullOrWhiteSpace($env:SELECTED_NEW_PCNAME)) {
    $hostUserList = Import-ModuleCsv -Path $hostCsvPath -RequiredColumns @("Enabled", "NewPCName", "UserName", "Password")
    if ($null -ne $hostUserList) {
        $pcUsers = @($hostUserList | Where-Object { $_.NewPCName -eq $env:SELECTED_NEW_PCNAME })
        if ($pcUsers.Count -gt 0) {
            Show-Info "Found $($pcUsers.Count) per-PC user(s) for '$($env:SELECTED_NEW_PCNAME)'"
            $userList = @($userList) + @($pcUsers)
        }
    }
}

if ($userList.Count -eq 0) {
    Show-Info "No users to delete"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No users to delete")
}

Show-Info "Loaded $($userList.Count) user definitions"
Write-Host ""

# ========================================
# List Users to Delete
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "User Deletion List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($user in $userList) {
    Write-Host "  UserName: $($user.UserName)" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Delete the users listed above?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# User Deletion Process
# ========================================
$successCount = 0
$skipCount = 0
$failCount = 0

foreach ($user in $userList) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Deleting User: $($user.UserName)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        # Check User Existence
        $existingUser = Get-LocalUser -Name $user.UserName -ErrorAction SilentlyContinue
        if (-not $existingUser) {
            Show-Skip "User '$($user.UserName)' does not exist"
            Write-Host ""
            $skipCount++
            continue
        }

        # Delete User
        Remove-LocalUser -Name $user.UserName -ErrorAction Stop
        Show-Success "Deleted user '$($user.UserName)'"
        $successCount++
    }
    catch {
        Show-Error "Failed to delete user '$($user.UserName)': $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Execution Results")