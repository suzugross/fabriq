# ========================================
# Built-in Administrator Configuration Script
# ========================================

$ADMIN_NAME = "Administrator"

Show-Info "Executing Built-in Administrator configuration..."
Write-Host ""

# ========================================
# Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "builtin_admin.csv"

$configList = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $configList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load builtin_admin.csv")
}
if ($configList.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

$config = $configList[0]
Write-Host ""

# ========================================
# Validate Password
# ========================================
if ([string]::IsNullOrWhiteSpace($config.Password)) {
    Show-Error "Password is empty in CSV. Password is required."
    return (New-ModuleResult -Status "Error" -Message "Password is empty")
}

# ========================================
# Verify Account Exists
# ========================================
$adminUser = Get-LocalUser -Name $ADMIN_NAME -ErrorAction SilentlyContinue
if ($null -eq $adminUser) {
    Show-Error "Account '$ADMIN_NAME' not found on this system"
    return (New-ModuleResult -Status "Error" -Message "Account '$ADMIN_NAME' not found")
}

# ========================================
# Display Configuration
# ========================================
$pwdExpireText = if ($config.PasswordNeverExpires -eq "1") { "Never" } else { "Expires" }

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Built-in Administrator Configuration" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""
Write-Host "  Target User:     $ADMIN_NAME" -ForegroundColor Yellow
Write-Host "  Current Status:  $(if ($adminUser.Enabled) { 'Enabled' } else { 'Disabled' })" -ForegroundColor Gray
Write-Host ""
Write-Host "  [Settings to Apply]" -ForegroundColor Cyan
Write-Host "    Account:            Enable" -ForegroundColor White
Write-Host "    Password:           ********" -ForegroundColor White
Write-Host "    Password Expiry:    $pwdExpireText" -ForegroundColor White
Write-Host "    Description:        $($config.Description)" -ForegroundColor White
Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above settings to '$ADMIN_NAME'?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Apply Configuration
# ========================================
try {
    # --- Enable account ---
    # Rows load with -FilterEnabled, so only Enabled=1 reaches here; this
    # module enables the built-in Administrator and sets its password.
    # (Enabled=0 rows are skipped - the standard run/skip semantics.)
    Show-Info "Enabling account..."
    Enable-LocalUser -Name $ADMIN_NAME -ErrorAction Stop
    Show-Success "Account enabled"

    # --- Password ---
    Show-Info "Setting password..."
    $securePassword = ConvertTo-SecureString $config.Password -AsPlainText -Force
    Set-LocalUser -Name $ADMIN_NAME -Password $securePassword -ErrorAction Stop
    Show-Success "Password set"

    # --- Password Never Expires ---
    Show-Info "Setting password expiry..."
    $pwdNeverExpires = ($config.PasswordNeverExpires -eq "1")
    Set-LocalUser -Name $ADMIN_NAME -PasswordNeverExpires $pwdNeverExpires -ErrorAction Stop
    Show-Success "Password expiry: $pwdExpireText"

    # --- Description ---
    if (-not [string]::IsNullOrWhiteSpace($config.Description)) {
        Show-Info "Setting description..."
        Set-LocalUser -Name $ADMIN_NAME -Description $config.Description -ErrorAction Stop
        Show-Success "Description set"
    }

    # ========================================
    # Step 5.5: Post-Apply Verification
    # ========================================
    # Read back the account state and confirm the applied settings took
    # effect. Partial verification: the password value cannot be read back
    # (the SAM hash is not retrievable), so it is intentionally NOT verified
    # - a presence check would false-PASS for a wrong password (same rationale
    # family as credential_config; see Guide).
    Write-Host ""
    Show-Info "Verifying applied settings..."
    Write-Host ""

    $verified = $true
    try {
        $check = Get-LocalUser -Name $ADMIN_NAME -ErrorAction Stop

        # (1) Account enabled
        if ($check.Enabled) {
            Write-Host "  [VERIFIED] Account enabled" -ForegroundColor Green
        } else {
            Write-Host "  [VERIFY FAILED] Account is not enabled" -ForegroundColor Red
            $verified = $false
        }

        # (2) Password-never-expires flag. PasswordExpires is $null when the
        #     account never expires. Edge: a system policy of
        #     MaximumPasswordAge=0 makes the 'expires' case read $null too -
        #     that is a false-FAIL (safe direction), never a false-PASS.
        $expectNeverExpires = ($config.PasswordNeverExpires -eq "1")
        $actualNeverExpires = ($null -eq $check.PasswordExpires)
        if ($actualNeverExpires -eq $expectNeverExpires) {
            Write-Host "  [VERIFIED] Password expiry: $pwdExpireText" -ForegroundColor Green
        } else {
            Write-Host "  [VERIFY FAILED] Password expiry mismatch (expected never-expires=$expectNeverExpires)" -ForegroundColor Red
            $verified = $false
        }

        # (3) Description (only verified when the apply step set it)
        if (-not [string]::IsNullOrWhiteSpace($config.Description)) {
            if ($check.Description -eq $config.Description) {
                Write-Host "  [VERIFIED] Description" -ForegroundColor Green
            } else {
                Write-Host "  [VERIFY FAILED] Description mismatch" -ForegroundColor Red
                $verified = $false
            }
        }
    }
    catch {
        # Fail-closed: if the read-back itself fails we cannot confirm.
        Write-Host "  [VERIFY FAILED] Could not read back account state: $_" -ForegroundColor Red
        $verified = $false
    }

    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Built-in Administrator configured successfully" -Verified $verified)
}
catch {
    Show-Error "Failed to configure Built-in Administrator: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Configuration failed: $_")
}
