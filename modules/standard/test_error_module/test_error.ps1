# ========================================
# Test Error Module
# ========================================
# Always fails with an error for testing error recovery flow.
# This module is for development/testing purposes only.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Test Error Module" -ForegroundColor Cyan
Show-Separator
Write-Host ""

Show-Info "This module will intentionally fail for testing purposes."
Write-Host ""

$cancelResult = Confirm-ModuleExecution -Message "Run test error module?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""
Show-Error "Simulated error: This module always fails."
Write-Host ""

return (New-ModuleResult -Status "Error" -Message "Simulated error for testing")
