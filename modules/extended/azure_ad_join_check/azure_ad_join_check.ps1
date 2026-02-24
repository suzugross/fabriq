# ========================================
# Azure AD (Entra ID) Join Check Script
# ========================================
# [PURPOSE]
# Run dsregcmd /status and check whether Azure AD Join is complete.
# If not yet joined, return Error to trigger retry from Script Looper.
#
# [NOTES]
# - Intended to be called from Script Looper (looper_list.csv).
# - Register with Condition=OnError to auto-retry until join completes.
# - No admin privileges required (dsregcmd runs as a normal user).
# ========================================

Write-Host ""
Show-Separator
Write-Host "Azure AD Join Check" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Prerequisite Check
# ========================================
# Verify dsregcmd.exe exists (should always exist on Windows 10+)
$dsregCmd = "$env:SystemRoot\System32\dsregcmd.exe"
if (-not (Test-Path $dsregCmd)) {
    Show-Error "dsregcmd.exe not found: $dsregCmd"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "dsregcmd.exe not found")
}

# ========================================
# Step 2: Execute dsregcmd /status
# ========================================
Show-Info "Checking Azure AD Join status..."
Write-Host ""

try {
    $dsregOutput = & dsregcmd /status 2>&1 | Out-String
}
catch {
    Show-Error "Failed to execute dsregcmd /status: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "dsregcmd execution failed: $_")
}

# ========================================
# Step 3: Parse and Display Status
# ========================================
# Example output line: "             AzureAdJoined : YES"
$isJoined = $dsregOutput -match "AzureAdJoined\s*:\s*YES"

# Extract current AzureAdJoined value for display
$currentState = ""
if ($dsregOutput -match "(AzureAdJoined\s*:\s*\w+)") {
    $currentState = $Matches[1].Trim()
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Azure AD Join Status" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($currentState) {
    Write-Host "  $currentState" -ForegroundColor White
} else {
    Write-Host "  AzureAdJoined : (not found in output)" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: Return Result
# ========================================
if ($isJoined) {
    Show-Success "Azure AD Join is complete."
    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Azure AD Joined: YES")
}

if ($currentState) {
    Show-Warning "Azure AD Join is not complete yet. ($currentState)"
} else {
    Show-Warning "Azure AD Join is not complete yet. (AzureAdJoined field not found in dsregcmd output)"
}

Write-Host ""
return (New-ModuleResult -Status "Error" -Message "Azure AD Join is not complete yet")
