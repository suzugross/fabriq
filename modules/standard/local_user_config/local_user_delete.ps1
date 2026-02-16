# ========================================
# Local User Deletion Script
# ========================================

Show-Info "Executing local user deletion..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "local_user_list.csv"

$userList = Import-CsvSafe -Path $csvPath -Description "local_user_list.csv"
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
if (-not (Confirm-Execution -Message "Delete the users listed above?")) {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

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
Show-Separator
Write-Host "Execution Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Skipped: $skipCount items" -ForegroundColor Yellow
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Show-Separator
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")