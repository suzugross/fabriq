# ========================================
# Local User Creation Script
# ========================================

Show-Info "Executing local user creation..."
Write-Host ""

function Get-LocalUserCreatePlan {
    # Classifies one CSV row before anything touches the machine. Pure (reads
    # no machine state, prints nothing) so the decision table is testable:
    # returns @{ Mode; Error }, a non-null Error meaning "refuse this row".
    #   Mode 'Password'   - the Password cell has a value
    #   Mode 'NoPassword' - the Password cell is there but empty
    # A missing Password FIELD ($null: the CSV line ends before that column)
    # is refused rather than read as "empty" - a truncated line must not turn
    # into an account without a password.
    param(
        [object]$Row
    )

    if ([string]::IsNullOrWhiteSpace("$($Row.UserName)")) {
        return @{ Mode = $null; Error = "UserName is empty" }
    }
    if ($null -eq $Row.Password) {
        return @{ Mode = $null; Error = "Password field is missing (the CSV line ends before the Password column)" }
    }
    if ("$($Row.Password)".Length -eq 0) {
        return @{ Mode = "NoPassword"; Error = $null }
    }

    return @{ Mode = "Password"; Error = $null }
}

# ========================================
# Load CSV
# ========================================
$csvPath = Resolve-ModuleDataPath -Path (Join-Path $PSScriptRoot "local_user_list.csv")

$userList = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "UserName", "Password")
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
$hostCsvPath = Resolve-ModuleDataPath -Path (Join-Path $PSScriptRoot "local_user_host_list.csv")

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

$enabledUsers = @($userList | Where-Object { $_.Enabled -eq "1" })
$disabledUsers = @($userList | Where-Object { $_.Enabled -ne "1" })

if ($enabledUsers.Count -eq 0) {
    Show-Info "No enabled users to create"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled users")
}
Write-Host ""

# ========================================
# List User Settings
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "User Creation List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($user in $enabledUsers) {
    $plan = Get-LocalUserCreatePlan -Row $user
    $source = if ($user.PSObject.Properties['NewPCName'] -and $user.NewPCName) { " [PC: $($user.NewPCName)]" } else { "" }
    $displayName = if ($user.Description) { "$($user.UserName) ($($user.Description))$source" } else { "$($user.UserName)$source" }
    $pwdExpire = if ($user.PasswordNeverExpires -eq "1") { "Never" } else { "Expires" }
    $pwdChange = if ($user.UserMayNotChangePassword -eq "1") { "Denied" } else { "Allowed" }
    Write-Host "  UserName: $displayName" -ForegroundColor Yellow
    if ($plan.Error) {
        Write-Host "    [INVALID] $($plan.Error) - this row will fail" -ForegroundColor Red
        Write-Host ""
        continue
    }
    if ($plan.Mode -eq "NoPassword") {
        Write-Host "    Password: (none) - the account will have NO password" -ForegroundColor Magenta
    } else {
        Write-Host "    Password: (set)"
    }
    Write-Host "    Password Expire: $pwdExpire / Change: $pwdChange / Group: $($user.Group)"
    Write-Host ""
}

foreach ($user in $disabledUsers) {
    $displayName = if ($user.Description) { "$($user.UserName) ($($user.Description))" } else { $user.UserName }
    Write-Host "  [DISABLED] $displayName" -ForegroundColor DarkGray
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Create users with above settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# User Creation Process
# ========================================
$successCount = 0
$failCount = 0

foreach ($user in $enabledUsers) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Creating User: $($user.UserName)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- Validate Row ---
    # Refused before the existence check, so a broken row fails the same way
    # whatever accounts the machine already has.
    $plan = Get-LocalUserCreatePlan -Row $user
    if ($plan.Error) {
        Show-Error "Row rejected (UserName '$($user.UserName)'): $($plan.Error)"
        Write-Host ""
        $failCount++
        continue
    }

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
        if ($plan.Mode -eq "NoPassword") {
            # An empty SecureString sets a blank password. New-LocalUser
            # -NoPassword is deliberately not used: it leaves "user must
            # change password at next logon" set and has no
            # -PasswordNeverExpires in its parameter set.
            Show-Warning "Creating user '$($user.UserName)' WITHOUT a password (Password cell is empty)"
            $securePassword = New-Object System.Security.SecureString
        } else {
            $securePassword = ConvertTo-SecureString $user.Password -AsPlainText -Force
        }
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
        if ($plan.Mode -eq "NoPassword") {
            Show-Success "Created user '$($user.UserName)' (no password)"
        } else {
            Show-Success "Created user '$($user.UserName)'"
        }
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
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying created users..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($user in $enabledUsers) {
    $displayName = $user.UserName

    # Verify user exists
    $localUser = Get-LocalUser -Name $user.UserName -ErrorAction SilentlyContinue
    if ($null -eq $localUser) {
        Write-Host "  [VERIFY FAILED] $displayName (user not found)" -ForegroundColor Red
        $verifyFail++
        continue
    }

    # Verify group membership
    $groupFail = $false
    if (-not [string]::IsNullOrWhiteSpace($user.Group)) {
        $groups = $user.Group -split ';'
        foreach ($group in $groups) {
            $group = $group.Trim()
            if ([string]::IsNullOrWhiteSpace($group)) { continue }

            $isMember = $false
            try {
                $members = @(Get-LocalGroupMember -Group $group -ErrorAction Stop)
                $isMember = $members | Where-Object { $_.Name -match "\\$($user.UserName)$" }
            }
            catch { }

            if (-not $isMember) {
                Write-Host "  [VERIFY FAILED] $displayName (not in group: $group)" -ForegroundColor Red
                $groupFail = $true
            }
        }
    }

    if ($groupFail) {
        $verifyFail++
    } else {
        Write-Host "  [VERIFIED] $displayName" -ForegroundColor Green
        $verifyPass++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)

return (New-BatchResult -Success $successCount -Fail $failCount -Verified $verified)