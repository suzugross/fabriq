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

# ========================================
# 1. Load CSV
# ========================================
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

# ========================================
# 2. Determine Target User
# ========================================
$targetEntry = $null

if (-not [string]::IsNullOrWhiteSpace($env:FABRIQ_AUTOLOGON_NO)) {
    # Profile mode: __AUTO_to_xxx__ specified
    $targetNo = $env:FABRIQ_AUTOLOGON_NO
    $targetEntry = $enabledEntries | Where-Object { $_.No -eq $targetNo } | Select-Object -First 1

    if ($null -eq $targetEntry) {
        Show-Error "No enabled entry found with No='$targetNo' in autologon_list.csv"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Entry No='$targetNo' not found or disabled")
    }

    Show-Info "Target specified from Profile: No=$targetNo"
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

Write-Host ""

# ========================================
# 3. Display Settings (Password Masked)
# ========================================
$displayUser = $targetEntry.User
$displayDomain = if ($targetEntry.Domain) { $targetEntry.Domain } else { "(Local)" }
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
# 4. Idempotency Check
# ========================================
try {
    $currentAutoLogon = Get-ItemProperty -Path $WINLOGON_PATH -Name "AutoAdminLogon" -ErrorAction SilentlyContinue
    $currentUser = Get-ItemProperty -Path $WINLOGON_PATH -Name "DefaultUserName" -ErrorAction SilentlyContinue
    $currentCount = Get-ItemProperty -Path $WINLOGON_PATH -Name "AutoLogonCount" -ErrorAction SilentlyContinue

    if ($currentAutoLogon.AutoAdminLogon -eq "1" -and
        $currentUser.DefaultUserName -eq $targetEntry.User -and
        $currentCount.AutoLogonCount -ge 1) {
        Show-Skip "AutoLogon already configured for '$displayUser' (Count=$($currentCount.AutoLogonCount))"
        Write-Host ""
        return (New-ModuleResult -Status "Skipped" -Message "Already configured for $displayUser")
    }
}
catch {
    # Registry read failed - proceed with configuration
}

# ========================================
# 5. Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Configure one-time AutoLogon for '$displayUser'?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# 6. Set Registry Values
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
# 7. Result
# ========================================
if ($failCount -gt 0) {
    Show-Warning "AutoLogon configuration completed with $failCount error(s)"
    Write-Host ""
    return (New-ModuleResult -Status "Partial" -Message "AutoLogon for $displayUser ($failCount errors)")
}

Show-Success "One-time AutoLogon configured for '$displayUser'"
Show-Info "AutoLogon will be cleared automatically after the next logon"
Write-Host ""
return (New-ModuleResult -Status "Success" -Message "AutoLogon configured for $displayUser (one-time)")
