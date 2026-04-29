# ========================================
# [MODULE NAME] Script
# ========================================
# [PURPOSE]
# One-line description of what this module does.
#
# [NOTES]
# - List any prerequisites or caveats here.
# - Examples: requires admin / requires network connectivity
# ========================================

Write-Host ""
Show-Separator
Write-Host "[MODULE NAME]" -ForegroundColor Cyan   # <-- replace with module display name
Show-Separator
Write-Host ""

# ========================================
# [OPTIONAL] use only when P/Invoke is needed
# Keep this block when calling into the Win32 API.
# Otherwise delete the entire block.
# ========================================
# Add-Type -TypeDefinition @'
# using System;
# using System.Runtime.InteropServices;
#
# public class TemplateHandler {
#     [DllImport("user32.dll")]
#     public static extern int SomeApiFunction(int param);
# }
# '@ -ErrorAction SilentlyContinue


# ========================================
# Step 1: Load CSV
# ========================================
# Resolve the CSV path relative to $PSScriptRoot.
# -RequiredColumns lists every column the script must reference.
# Add any extra columns beyond Enabled / Description here.
#
# [Segment support]
# Adding a Segment column to the CSV lets the Profile invoke this
# module with a segment filter; rows are auto-filtered by
# Import-ModuleCsv with no module-side code change needed.
# ========================================
$csvPath = Join-Path $PSScriptRoot "_template_list.csv"   # <-- rename to your actual CSV file

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "TargetName")           # <-- match your actual CSV column names

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load _template_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Prerequisite check (early return)
# ========================================
# Verify that resources required by this script (directories,
# executables, etc.) exist. Return immediately on missing
# prerequisites so failures surface before any side effect.
# Delete the entire block when no prerequisite check is needed.
# ========================================
# Example: working directory existence check
# $workDir = Join-Path $PSScriptRoot "files"
# if (-not (Test-Path $workDir)) {
#     Show-Error "'files' directory not found: $workDir"
#     Write-Host ""
#     return (New-ModuleResult -Status "Error" -Message "'files' directory not found")
# }


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
# Tell the user what is about to happen.
# Make the WHAT and the EXPECTED OUTCOME obvious at a glance.
# When probing current state, color-code with labels like
# [APPLY] / [SKIP] / [NOT FOUND].
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Items" -ForegroundColor Yellow    # <-- change the heading to fit your module
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetName }

    # Per-item current-state probing goes here.
    # Example: file existence check, read of current setting value, etc.

    Write-Host "  [APPLY] $displayName" -ForegroundColor Yellow   # <-- swap color and label based on probed state
    Write-Host "    Target: $($item.TargetName)" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: User confirmation
# ========================================
# Confirm-ModuleExecution prompts the user with Y/N and returns a
# Cancelled ModuleResult when the user answers N.
# AutoPilot mode treats every prompt as auto-Y.
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above settings?"   # <-- replace with a confirmation message that fits your action
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Apply-settings loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetName }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan   # <-- replace with a verb that fits the action
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Per-item prerequisite check (Skip path)
    # ----------------------------------------
    # Run BEFORE try{}. On unmet condition, call Show-Skip and
    # continue so the loop tally counts it as a skip.
    # ----------------------------------------
    # Example: file existence check
    # if (-not (Test-Path $somePath)) {
    #     Show-Skip "File not found: $somePath"
    #     Write-Host ""
    #     $skipCount++
    #     continue
    # }

    # ----------------------------------------
    # Main work
    # ----------------------------------------
    try {
        # Put the actual implementation here.
        # Example: registry write, file copy, API call, etc.

        # On success
        Show-Success "Completed: $displayName"
        $successCount++
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 5.5: Post-Apply Verification (Optional)
# ========================================
# If this module can verify that settings were applied correctly,
# implement verification logic here.
# Read back the actual system state and compare with expected values.
#
# To enable verification:
# 1. Uncomment the block below
# 2. Replace the comparison logic with module-specific checks
# 3. Pass -Verified $verified to New-BatchResult in Step 6
#
# Reference implementations:
#   - reg_hklm_config : Uses Test-RegistryValueMatch to verify registry values
#   - firewall_config : Uses Get-NetFirewallProfile to verify profile states
#   - hostname_config : Checks pending hostname in registry
# ========================================
# $verifyPass = 0
# $verifyFail = 0
#
# foreach ($item in $enabledItems) {
#     $displayName = if ($item.Description) { $item.Description } else { $item.TargetName }
#
#     # TODO: Read back the current state
#     # $actual = ...
#     # $expected = $item.TargetName  # or relevant column
#
#     # if ($actual -eq $expected) {
#     #     Write-Host "  [VERIFIED] $displayName" -ForegroundColor Green
#     #     $verifyPass++
#     # } else {
#     #     Write-Host "  [VERIFY FAILED] $displayName (expected: $expected, actual: $actual)" -ForegroundColor Red
#     #     $verifyFail++
#     # }
# }
#
# $verified = ($verifyFail -eq 0)


# ========================================
# Step 6: Aggregate and return result
# ========================================
# New-BatchResult tallies the Success/Skip/Fail counts and returns a
# New-ModuleResult whose Status (Success / Partial / Error / Skipped)
# is auto-decided based on those counts.
# If Step 5.5 is implemented, add: -Verified $verified
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "[MODULE NAME] Results")   # <-- replace title with your module's display name
