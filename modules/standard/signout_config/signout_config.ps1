# ========================================
# Sign-Out Script
# ========================================
# Signs out the current user session.
# fabriq process will be terminated after
# sign-out. All remaining modules in the
# profile will NOT execute.
#
# NOTES:
# - Place this module LAST in the profile.
# - AutoPilot mode auto-confirms and signs
#   out without user interaction.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Sign-Out" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Warning Display
# ========================================
Write-Host "----------------------------------------" -ForegroundColor Red
Write-Host "  SIGN-OUT REQUESTED" -ForegroundColor Red
Write-Host "----------------------------------------" -ForegroundColor Red
Write-Host ""
Show-Warning "fabriq will be TERMINATED after sign-out."
Show-Warning "All remaining modules in this profile will NOT execute."
Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor Red
Write-Host ""


# ========================================
# Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Sign out now?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Pre-signout: save result before process dies
# ========================================
# logoff.exe terminates this process immediately.
# New-ModuleResult sets $global:_LastModuleResult as a fallback
# so the framework can capture the result if possible.
$result = New-ModuleResult -Status "Success" -Message "Sign-out initiated"

Show-Warning "Signing out. fabriq process will be terminated."
Write-Host ""


# ========================================
# Countdown and Sign-Out
# ========================================
Invoke-CountdownSignout -Seconds 7

return $result
