# ========================================
# Local User Creation Script
# ========================================

Write-Host "Executing local user creation..." -ForegroundColor Cyan
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
    Write-Host "[INFO] Canceled" -ForegroundColor Yellow
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
            Write-Host "[SKIP] User '$($user.UserName)' already exists" -ForegroundColor Yellow
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
        Write-Host "[SUCCESS] Created user '$($user.UserName)'" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Failed to create user '$($user.UserName)': $_" -ForegroundColor Red
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
                Write-Host "[SUCCESS] Added to group '$group'" -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Failed to add to group '$group': $_" -ForegroundColor Red
            }
        }
    }

    $successCount++
    Write-Host ""
}

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Execution Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Success: $successCount items" -ForegroundColor Green
Write-Host "  Failed: $failCount items" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Fail: $failCount")