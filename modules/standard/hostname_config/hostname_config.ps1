# ========================================
# Hostname Change Script
# ========================================

Show-Info "Executing hostname change..."
Write-Host ""

# ========================================
# Get Target Hostname from Environment
# ========================================
$newHostname = $env:SELECTED_NEW_PCNAME

if ([string]::IsNullOrWhiteSpace($newHostname)) {
    Show-Skip "NewPCName is not specified in hostlist. Hostname change skipped."
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "NewPCName not specified in hostlist")
}

$currentHostname = $env:COMPUTERNAME

# ========================================
# Display Change Info
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Hostname Change" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""
Write-Host "  Current Hostname:  $currentHostname" -ForegroundColor White
Write-Host "  New Hostname:      $newHostname" -ForegroundColor Yellow
Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# Idempotency Check
# ========================================
if ($currentHostname -eq $newHostname) {
    Show-Skip "Current hostname is already '$newHostname'. No change needed."
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Current hostname is already $newHostname")
}

# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Change hostname: $currentHostname -> $newHostname ?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Change Hostname
# ========================================
try {
    Rename-Computer -NewName $newHostname -Force -ErrorAction Stop
    Show-Success "Hostname changed: $currentHostname -> $newHostname"
}
catch {
    Show-Error "Failed to change hostname: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to change hostname: $_")
}

Write-Host ""

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
# Hostname change requires restart to take effect.
# Verify that the pending hostname in registry matches the expected value.
$verified = $null
try {
    $pendingName = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -Name "ComputerName" -ErrorAction Stop).ComputerName
    $verified = ($pendingName -eq $newHostname)

    if ($verified) {
        Show-Success "[VERIFIED] Pending hostname: $pendingName"
    } else {
        Show-Warning "[VERIFY FAILED] Expected: $newHostname, Pending: $pendingName"
    }
}
catch {
    Show-Warning "Could not verify pending hostname: $_"
}

Write-Host ""

Show-Warning "Restart is required to apply the hostname change."
Write-Host ""

return (New-ModuleResult -Status "Success" -Message "Hostname changed: $currentHostname -> $newHostname (restart required)" -Verified $verified)
