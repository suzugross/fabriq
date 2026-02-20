# ========================================
# User Profile Deletion Script
# ========================================
# Deletes user profile folders specified in
# profile_list.csv. Uses Win32_UserProfile
# (WMI) to properly unregister profiles from
# Windows, with a physical folder removal
# fallback for orphaned profile directories.
#
# NOTES:
# - Administrator privileges required.
# - Profiles in use (active login) cannot be
#   deleted and will be reported as Error.
# ========================================

# Check Administrator Privileges
if (-not (Test-AdminPrivilege)) {
    Show-Error "This script requires administrator privileges."
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

Write-Host ""
Show-Separator
Write-Host "User Profile Deletion" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV load
# ========================================
$csvPath = Join-Path $PSScriptRoot "profile_list.csv"

$profileList = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "UserName")

if ($null -eq $profileList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load profile_list.csv")
}
if ($profileList.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries in profile_list.csv")
}


# ========================================
# Step 2: Pre-flight check
# ========================================
$usersBase = Join-Path $env:SystemDrive "Users"

if (-not (Test-Path $usersBase)) {
    Show-Error "Users directory not found: $usersBase"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Users directory not found: $usersBase")
}


# ========================================
# Step 3: Pre-execution display
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Profiles to Delete" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($entry in $profileList) {
    $displayName = if ($entry.Description) { $entry.Description } else { $entry.UserName }
    $profilePath = Join-Path $usersBase $entry.UserName

    if (Test-Path $profilePath) {
        Write-Host "  [APPLY]  $displayName" -ForegroundColor Yellow
        Write-Host "    Path: $profilePath" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  [SKIP]   $displayName" -ForegroundColor DarkGray
        Write-Host "    Path: $profilePath  (not found)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Delete the listed user profiles?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Deletion loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($entry in $profileList) {
    $displayName = if ($entry.Description) { $entry.Description } else { $entry.UserName }
    $userName    = $entry.UserName
    $profilePath = Join-Path $usersBase $userName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Deleting Profile: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- Existence check (outside try) ---
    if (-not (Test-Path $profilePath)) {
        Show-Skip "Profile folder not found: $profilePath"
        Write-Host ""
        $skipCount++
        continue
    }

    Show-Info "Path: $profilePath"

    try {
        # Stage 1: WMI deletion (properly unregisters the profile from Windows)
        $wmiProfile = Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue |
            Where-Object { $_.LocalPath -like "*\$userName" }

        if ($wmiProfile) {
            $wmiProfile | Remove-CimInstance -ErrorAction Stop
            Show-Info "WMI profile record removed"
        }
        else {
            Show-Info "No WMI record found (orphaned folder) — proceeding to folder removal"
        }

        # Stage 2: Physical folder removal fallback
        # (WMI deletion does not always remove the folder)
        if (Test-Path $profilePath) {
            Remove-Item -Path $profilePath -Recurse -Force -ErrorAction Stop
            Show-Info "Profile folder removed"
        }

        Show-Success "Deleted profile: $displayName"
        $successCount++
    }
    catch {
        Show-Error "Failed to delete profile '$displayName': $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Profile Deletion Results")
