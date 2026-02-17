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
$enableText = if ($config.Enabled -eq "1") { "Enable" } else { "Disable" }
$pwdExpireText = if ($config.PasswordNeverExpires -eq "1") { "Never" } else { "Expires" }

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Built-in Administrator Configuration" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""
Write-Host "  Target User:     $ADMIN_NAME" -ForegroundColor Yellow
Write-Host "  Current Status:  $(if ($adminUser.Enabled) { 'Enabled' } else { 'Disabled' })" -ForegroundColor Gray
Write-Host ""
Write-Host "  [Settings to Apply]" -ForegroundColor Cyan
Write-Host "    Account:            $enableText" -ForegroundColor White
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
    # --- Enable / Disable ---
    Show-Info "Setting account status..."
    if ($config.Enabled -eq "1") {
        Enable-LocalUser -Name $ADMIN_NAME -ErrorAction Stop
        Show-Success "Account enabled"
    }
    else {
        Disable-LocalUser -Name $ADMIN_NAME -ErrorAction Stop
        Show-Success "Account disabled"
    }

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

    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Built-in Administrator configured successfully")
}
catch {
    Show-Error "Failed to configure Built-in Administrator: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Configuration failed: $_")
}
