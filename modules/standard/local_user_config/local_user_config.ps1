# ========================================
# Local User Creation Script
# ========================================

Show-Info "Executing local user creation..."
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
# List User Settings
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "User Creation List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($user in $userList) {
    $pwdExpire = if ($user.PasswordNeverExpires -eq "1") { "Never" } else { "Expires" }
    $pwdChange = if ($user.UserMayNotChangePassword -eq "1") { "Denied" } else { "Allowed" }
    Write-Host "  UserName: $($user.UserName)" -ForegroundColor Yellow
    Write-Host "    Password Expire: $pwdExpire / Change: $pwdChange / Group: $($user.Group)"
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
if (-not (Confirm-Execution -Message "Create users with above settings?")) {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# User Creation Process
# ========================================
$successCount = 0
$failCount = 0

foreach ($user in $userList) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Creating User: $($user.UserName)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- Create User ---
    try {
        # Check existing user
        $existingUser = Get-LocalUser -Name $user.UserName -ErrorAction SilentlyContinue
        if ($existingUser) {
            Show-Skip "User '$($user.UserName)' already exists"
            Write-Host ""
            continue
        }

        # Build parameters
        $securePassword = ConvertTo-SecureString $user.Password -AsPlainText -Force
        $params = @{
            Name                    = $user.UserName
            Password                = $securePassword
            AccountNeverExpires     = $true
        }

        if ($user.PasswordNeverExpires -eq "1") {
            $params.PasswordNeverExpires = $true
        }

        if ($user.UserMayNotChangePassword -eq "1") {
            $params.UserMayNotChangePassword = $true
        }

        # Create user
        New-LocalUser @params | Out-Null
        Show-Success "Created user '$($user.UserName)'"
    }
    catch {
        Show-Error "Failed to create user '$($user.UserName)': $_"
        Write-Host ""
        $failCount++
        continue
    }

    # --- Add to Group ---
    if (-not [string]::IsNullOrWhiteSpace($user.Group)) {
        $groups = $user.Group -split ';'
        foreach ($group in $groups) {
            $group = $group.Trim()
            if ([string]::IsNullOrWhiteSpace($group)) { continue }

            try {
                Add-LocalGroupMember -Group $group -Member $user.UserName -ErrorAction Stop
                Show-Success "Added to group '$group'"
            }
            catch {
                Show-Error "Failed to add to group '$group': $_"
            }
        }
    }

    $successCount++
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Show-Separator
Write-Host "Execution Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Show-Separator
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Fail: $failCount")