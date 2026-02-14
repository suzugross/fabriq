# ========================================
# Local User Deletion Script
# ========================================

Write-Host "Executing local user deletion..." -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "local_user_list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] local_user_list.csv not found: $csvPath" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "local_user_list.csv not found")
}

try {
    $userList = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load local_user_list.csv: $_" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load local_user_list.csv: $_")
}

if ($userList.Count -eq 0) {
    Write-Host "[ERROR] local_user_list.csv contains no data" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "local_user_list.csv contains no data")
}

Write-Host "[INFO] Loaded $($userList.Count) user definitions" -ForegroundColor Cyan
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
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
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
            Write-Host "[SKIP] User '$($user.UserName)' does not exist" -ForegroundColor Yellow
            Write-Host ""
            $skipCount++
            continue
        }

        # Delete User
        Remove-LocalUser -Name $user.UserName -ErrorAction Stop
        Write-Host "[SUCCESS] Deleted user '$($user.UserName)'" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to delete user '$($user.UserName)': $_" -ForegroundColor Red
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
Write-Host "  Skipped: $skipCount items" -ForegroundColor Yellow
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")