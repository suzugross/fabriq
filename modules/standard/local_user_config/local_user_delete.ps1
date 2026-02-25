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
if ($null -eq $userList -or $userList.Count -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load local_user_list.csv")
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