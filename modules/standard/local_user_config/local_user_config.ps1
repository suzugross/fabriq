# ========================================
# Local User Creation Script
# ========================================

Show-Info "Executing local user creation..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "local_user_list.csv"

$userList = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "UserName", "Password")
if ($null -eq $userList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load local_user_list.csv")
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

$enabledUsers = @($userList | Where-Object { $_.Enabled -eq "1" })
$disabledUsers = @($userList | Where-Object { $_.Enabled -ne "1" })

if ($enabledUsers.Count -eq 0) {
    Show-Info "No enabled users in local_user_list.csv"
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
    $source = if ($user.PSObject.Properties['NewPCName'] -and $user.NewPCName) { " [PC: $($user.NewPCName)]" } else { "" }
    $displayName = if ($user.Description) { "$($user.UserName) ($($user.Description))$source" } else { "$($user.UserName)$source" }
    $pwdExpire = if ($user.PasswordNeverExpires -eq "1") { "Never" } else { "Expires" }
    $pwdChange = if ($user.UserMayNotChangePassword -eq "1") { "Denied" } else { "Allowed" }
    Write-Host "  UserName: $displayName" -ForegroundColor Yellow
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

return (New-BatchResult -Success $successCount -Fail $failCount)