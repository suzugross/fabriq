# ========================================
# One-time AutoLogon Configuration
# ========================================
# Sets Windows AutoLogon registry values for
# one-time automatic logon after restart.
# Uses AutoLogonCount=1 so Windows automatically
# clears credentials after the single logon.
# ========================================

Write-Host ""
Show-Separator
Write-Host "One-time AutoLogon Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

$WINLOGON_PATH = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"

# Profile token standing in for "the local user this PC gets" instead of a
# literal user name. Brackets are invalid in local account names, so the
# token can never collide with a real entry in autologon_list.csv.
$HOST_USER_TOKEN = "[HOSTUSER]"

function Get-HostAutoLogonTarget {
    # Picks the single local user to log on to out of local_user_config's
    # per-PC rows. Pure (reads no machine state, prints nothing) so the
    # decision table is testable: returns @{ UserName; Password; Error },
    # a non-null Error meaning "refuse". Every ambiguity is an Error - an
    # operator on site cannot be asked which account the customer meant.
    param(
        [object[]]$Rows,
        [string]$PcName
    )

    if ([string]::IsNullOrWhiteSpace($PcName)) {
        return @{ UserName = $null; Password = $null
                  Error = "No host is selected (SELECTED_NEW_PCNAME is empty), so the per-PC user cannot be resolved" }
    }

    $mine = @($Rows | Where-Object { $_.NewPCName -eq $PcName -and $_.Enabled -eq "1" })
    if ($mine.Count -eq 0) {
        return @{ UserName = $null; Password = $null
                  Error = "No enabled user row for PC '$PcName' in local_user_host_list.csv" }
    }

    # Collapse repeated rows for one user (Windows account names are
    # case-insensitive). Rows that disagree on the password cannot be
    # collapsed - which one the live account carries is unknowable here.
    $users = @()
    foreach ($row in $mine) {
        $name = "$($row.UserName)".Trim()
        $flagged = $false
        if ($row.PSObject.Properties['AutoLogon']) {
            $flagged = ("$($row.AutoLogon)".Trim() -eq "1")
        }

        $existing = $users | Where-Object { $_.UserName -ieq $name } | Select-Object -First 1
        if ($existing) {
            if ($existing.Password -cne "$($row.Password)") {
                return @{ UserName = $null; Password = $null
                          Error = "User '$name' appears more than once for PC '$PcName' with different passwords" }
            }
            if ($flagged) { $existing.Flagged = $true }
            continue
        }
        $users += [PSCustomObject]@{ UserName = $name; Password = "$($row.Password)"; Flagged = $flagged }
    }

    # AutoLogon=1 elects the target among several users of THIS PC. Flags on
    # other PCs' rows are already out of the candidate set (filtered by
    # NewPCName above), so one flag per PC is the normal way to fill the CSV.
    $flaggedUsers = @($users | Where-Object { $_.Flagged })
    if ($flaggedUsers.Count -gt 1) {
        return @{ UserName = $null; Password = $null
                  Error = "AutoLogon=1 is set on $($flaggedUsers.Count) users of PC '$PcName' ($(($flaggedUsers.UserName) -join ', ')) - exactly one is required" }
    }
    if ($flaggedUsers.Count -eq 0 -and $users.Count -gt 1) {
        return @{ UserName = $null; Password = $null
                  Error = "PC '$PcName' has $($users.Count) enabled users ($(($users.UserName) -join ', ')) - mark one of them with AutoLogon=1" }
    }

    $target = if ($flaggedUsers.Count -eq 1) { $flaggedUsers[0] } else { $users[0] }

    if ([string]::IsNullOrEmpty($target.Password)) {
        return @{ UserName = $null; Password = $null
                  Error = "User '$($target.UserName)' has an empty password in local_user_host_list.csv" }
    }
    if ($target.Password.StartsWith("ENC:")) {
        # Import-ModuleCsv leaves the ciphertext in place when it cannot
        # decrypt; writing it to DefaultPassword would arm a logon that
        # fails at the next boot.
        return @{ UserName = $null; Password = $null
                  Error = "Password of '$($target.UserName)' is still encrypted (ENC:) - it could not be decrypted with the master passphrase" }
    }

    return @{ UserName = $target.UserName; Password = $target.Password; Error = $null }
}

# ========================================
# 1. Determine Target Entry
# ========================================
$targetEntry = $null
$requestedUser = "$($env:FABRIQ_AUTOLOGON_USER)".Trim()

if ($requestedUser -eq $HOST_USER_TOKEN) {
    # Host-bound mode (__AUTO_to_[HOSTUSER]__): the target is whichever local
    # user local_user_config creates on THIS PC, so neither the name nor the
    # password has to be restated in autologon_list.csv.
    Show-Info "Host-bound mode: resolving the local user of PC '$($env:SELECTED_NEW_PCNAME)'"

    $hostCsvPath = Resolve-ModuleDataPath -Path ([System.IO.Path]::GetFullPath(
        (Join-Path $PSScriptRoot "..\local_user_config\local_user_host_list.csv")))

    if (-not (Test-Path -LiteralPath $hostCsvPath)) {
        Show-Error "local_user_host_list.csv not found: $hostCsvPath"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "local_user_host_list.csv not found")
    }

    # Segment filtering and ENC: decryption happen in the kernel. An empty
    # result means "no row matched this Segment" and $null a load failure;
    # both land on the same Error below, which names the PC and the Segment.
    # NOT @(Import-ModuleCsv ...): the kernel returns ,@(rows), so wrapping
    # the call collapses every row into one element whose properties are
    # arrays (PS 5.1 comma-return trap).
    $hostRows = Import-ModuleCsv -Path $hostCsvPath `
        -RequiredColumns @("Enabled", "NewPCName", "UserName", "Password")
    if ($null -eq $hostRows) { $hostRows = @() }

    $selection = Get-HostAutoLogonTarget -Rows @($hostRows) -PcName $env:SELECTED_NEW_PCNAME
    if ($selection.Error) {
        Show-Error $selection.Error
        if (-not [string]::IsNullOrWhiteSpace($env:FABRIQ_SEGMENT)) {
            Show-Info "Active Segment is [$($env:FABRIQ_SEGMENT)] - only rows carrying that Segment are candidates"
        }
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message $selection.Error)
    }

    # The account must exist BEFORE the restart this autologon is for. Arming
    # a logon for a user local_user_config has not created yet parks the next
    # boot on the logon screen, where a profile resume cannot continue.
    $localUser = Get-LocalUser -Name $selection.UserName -ErrorAction SilentlyContinue
    if ($null -eq $localUser) {
        Show-Error "Local user '$($selection.UserName)' does not exist yet - place this row after local_user_config in the profile"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Local user '$($selection.UserName)' not created yet")
    }
    if (-not $localUser.Enabled) {
        Show-Error "Local user '$($selection.UserName)' is disabled - it cannot log on"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Local user '$($selection.UserName)' is disabled")
    }

    # Domain "." pins the logon to this machine's own SAM. Measured on the AD
    # lab (2026-09-16) on a domain-joined and on a workgroup PC: unlike a
    # baked computer name it survives a pending rename, and unlike an empty
    # value it stays unambiguous when the customer's domain holds an account
    # of the same name.
    $targetEntry = [PSCustomObject]@{
        No          = "-"
        User        = $selection.UserName
        Password    = $selection.Password
        Domain      = "."
        Description = "per-PC local user of $($env:SELECTED_NEW_PCNAME)"
    }
}
else {
    # ----------------------------------------
    # 1a. Load CSV
    # ----------------------------------------
    $csvPath = Join-Path $PSScriptRoot "autologon_list.csv"

    $allEntries = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "No", "User", "Password")
    if ($null -eq $allEntries) {
        return (New-ModuleResult -Status "Error" -Message "Failed to load autologon_list.csv")
    }

    $enabledEntries = @($allEntries | Where-Object { $_.Enabled -eq "1" })

    if ($enabledEntries.Count -eq 0) {
        Show-Info "No enabled entries in autologon_list.csv"
        Write-Host ""
        return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
    }

    # ----------------------------------------
    # 1b. Determine Target User
    # ----------------------------------------
    if (-not [string]::IsNullOrWhiteSpace($requestedUser)) {
        # Profile mode: __AUTO_to_xxx__ specified
        $targetUser = $requestedUser
        $targetEntry = $enabledEntries | Where-Object { $_.User -eq $targetUser } | Select-Object -First 1

        if ($null -eq $targetEntry) {
            Show-Error "No enabled entry found with User='$targetUser' in autologon_list.csv"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Entry User='$targetUser' not found or disabled")
        }

        Show-Info "Target specified from Profile: User=$targetUser"
    }
    else {
        # Script Menu mode: manual selection or first entry
        if ($enabledEntries.Count -eq 1) {
            $targetEntry = $enabledEntries[0]
        }
        else {
            Write-Host "Available users:" -ForegroundColor Cyan
            Write-Host ""
            foreach ($entry in $enabledEntries) {
                $desc = if ($entry.Description) { " - $($entry.Description)" } else { "" }
                $domainInfo = if ($entry.Domain) { " ($($entry.Domain))" } else { " (Local)" }
                Write-Host "  [$($entry.No)] $($entry.User)$domainInfo$desc" -ForegroundColor White
            }
            Write-Host ""

            while ($true) {
                Write-Host -NoNewline "Select No: "
                $userInput = Read-Host
                $targetEntry = $enabledEntries | Where-Object { $_.No -eq $userInput } | Select-Object -First 1
                if ($targetEntry) { break }
                Show-Error "Invalid No. Please try again."
            }
        }
    }
}

Write-Host ""

# ========================================
# 2. Display Settings (Password Masked)
# ========================================
$displayUser = $targetEntry.User
$displayDomain = if ([string]::IsNullOrWhiteSpace($targetEntry.Domain)) { "(Local)" }
                 elseif ($targetEntry.Domain -eq ".") { ". (local account on this PC)" }
                 else { $targetEntry.Domain }
$displayDesc = if ($targetEntry.Description) { $targetEntry.Description } else { "-" }
$maskedPassword = if ($targetEntry.Password.Length -gt 2) {
    $targetEntry.Password[0] + ("*" * ($targetEntry.Password.Length - 2)) + $targetEntry.Password[-1]
} else {
    "**"
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "AutoLogon Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  No:          $($targetEntry.No)" -ForegroundColor White
Write-Host "  User:        $displayUser" -ForegroundColor White
Write-Host "  Password:    $maskedPassword" -ForegroundColor White
Write-Host "  Domain:      $displayDomain" -ForegroundColor White
Write-Host "  Description: $displayDesc" -ForegroundColor White
Write-Host "  Count:       1 (one-time)" -ForegroundColor White
Write-Host ""
Write-Host "  Registry:    $WINLOGON_PATH" -ForegroundColor Gray
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# 3. Idempotency Check
# ========================================
try {
    $currentAutoLogon = Get-ItemProperty -Path $WINLOGON_PATH -Name "AutoAdminLogon" -ErrorAction SilentlyContinue
    $currentUser = Get-ItemProperty -Path $WINLOGON_PATH -Name "DefaultUserName" -ErrorAction SilentlyContinue
    $currentCount = Get-ItemProperty -Path $WINLOGON_PATH -Name "AutoLogonCount" -ErrorAction SilentlyContinue
    # Password and domain must match too: skipping on a stale
    # DefaultPassword (CSV corrected after a first run) or a stale
    # DefaultDomainName (left by another configuration) ships an
    # autologon that fails at the next boot while reporting Skipped.
    $currentPass = Get-ItemProperty -Path $WINLOGON_PATH -Name "DefaultPassword" -ErrorAction SilentlyContinue
    $currentDomain = Get-ItemProperty -Path $WINLOGON_PATH -Name "DefaultDomainName" -ErrorAction SilentlyContinue
    $domainMatches = if ([string]::IsNullOrWhiteSpace($targetEntry.Domain)) {
        [string]::IsNullOrEmpty($currentDomain.DefaultDomainName)
    } else {
        $currentDomain.DefaultDomainName -eq $targetEntry.Domain
    }

    if ($currentAutoLogon.AutoAdminLogon -eq "1" -and
        $currentUser.DefaultUserName -eq $targetEntry.User -and
        $currentPass.DefaultPassword -ceq $targetEntry.Password -and
        $domainMatches -and
        $currentCount.AutoLogonCount -ge 1) {
        Show-Skip "AutoLogon already configured for '$displayUser' (Count=$($currentCount.AutoLogonCount))"
        Write-Host ""
        return (New-ModuleResult -Status "Skipped" -Message "Already configured for $displayUser" -Verified $true)
    }
}
catch {
    # Registry read failed - proceed with configuration
}

# ========================================
# 4. Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Configure one-time AutoLogon for '$displayUser'?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# 5. Set Registry Values
# ========================================
Show-Info "Configuring AutoLogon registry..."
Write-Host ""

$failCount = 0

# AutoAdminLogon
try {
    Set-ItemProperty -Path $WINLOGON_PATH -Name "AutoAdminLogon" -Value "1" -Type String -Force -ErrorAction Stop
    Show-Success "AutoAdminLogon = 1"
}
catch {
    Show-Error "Failed to set AutoAdminLogon: $_"
    $failCount++
}

# DefaultUserName
try {
    Set-ItemProperty -Path $WINLOGON_PATH -Name "DefaultUserName" -Value $targetEntry.User -Type String -Force -ErrorAction Stop
    Show-Success "DefaultUserName = $displayUser"
}
catch {
    Show-Error "Failed to set DefaultUserName: $_"
    $failCount++
}

# DefaultPassword
try {
    Set-ItemProperty -Path $WINLOGON_PATH -Name "DefaultPassword" -Value $targetEntry.Password -Type String -Force -ErrorAction Stop
    Show-Success "DefaultPassword = (set)"
}
catch {
    Show-Error "Failed to set DefaultPassword: $_"
    $failCount++
}

# DefaultDomainName
if (-not [string]::IsNullOrWhiteSpace($targetEntry.Domain)) {
    try {
        Set-ItemProperty -Path $WINLOGON_PATH -Name "DefaultDomainName" -Value $targetEntry.Domain -Type String -Force -ErrorAction Stop
        Show-Success "DefaultDomainName = $($targetEntry.Domain)"
    }
    catch {
        Show-Error "Failed to set DefaultDomainName: $_"
        $failCount++
    }
}
else {
    # Clear a stale DefaultDomainName left by another configuration -
    # with it present, a local-account autologon would try
    # STALEDOMAIN\user at the next boot and fail.
    Remove-ItemProperty -Path $WINLOGON_PATH -Name "DefaultDomainName" -Force -ErrorAction SilentlyContinue
    Show-Info "DefaultDomainName cleared (local account)"
}

# AutoLogonCount (DWORD = 1 for one-time logon)
try {
    Set-ItemProperty -Path $WINLOGON_PATH -Name "AutoLogonCount" -Value 1 -Type DWord -Force -ErrorAction Stop
    Show-Success "AutoLogonCount = 1"
}
catch {
    Show-Error "Failed to set AutoLogonCount: $_"
    $failCount++
}

Write-Host ""

# ========================================
# 5.5. Post-Apply Verification
# ========================================
$verified = $null
try {
    $v = Get-ItemProperty -Path $WINLOGON_PATH -ErrorAction Stop
    $vDomainOk = if ([string]::IsNullOrWhiteSpace($targetEntry.Domain)) {
        [string]::IsNullOrEmpty($v.DefaultDomainName)
    } else {
        $v.DefaultDomainName -eq $targetEntry.Domain
    }
    $verified = [bool]($v.AutoAdminLogon -eq "1" -and
                 $v.DefaultUserName -eq $targetEntry.User -and
                 $v.DefaultPassword -ceq $targetEntry.Password -and
                 $vDomainOk -and
                 $v.AutoLogonCount -ge 1)
    if ($verified) {
        Write-Host "  [VERIFIED] AutoLogon state matches the CSV entry" -ForegroundColor Green
    } else {
        Write-Host "  [VERIFY FAILED] AutoLogon state does not match the CSV entry" -ForegroundColor Red
    }
}
catch {
    $verified = $false
    Write-Host "  [VERIFY FAILED] Could not read Winlogon values: $_" -ForegroundColor Red
}
Write-Host ""

# ========================================
# 6. Result
# ========================================
if ($failCount -gt 0) {
    Show-Warning "AutoLogon configuration completed with $failCount error(s)"
    Write-Host ""
    return (New-ModuleResult -Status "Partial" -Message "AutoLogon for $displayUser ($failCount errors)" -Verified $verified)
}

Show-Success "One-time AutoLogon configured for '$displayUser'"
Show-Info "AutoLogon will be cleared automatically after the next logon"
Write-Host ""
return (New-ModuleResult -Status "Success" -Message "AutoLogon configured for $displayUser (one-time)" -Verified $verified)
