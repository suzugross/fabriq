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
    Show-Error "No host selected. Please select a host from the main menu first."
    return (New-ModuleResult -Status "Error" -Message "No host selected (SELECTED_NEW_PCNAME is empty)")
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

Show-Warning "Restart is required to apply the hostname change."
Write-Host ""

return (New-ModuleResult -Status "Success" -Message "Hostname changed: $currentHostname -> $newHostname (restart required)")
