# ========================================
# File Delete - CSV-Driven File/Folder Deletion
# ========================================
# Deletes files and folders listed in delete_list.csv.
# Supports environment variable expansion and per-entry
# missing-file behavior (Skip or Error).
# ========================================

Write-Host ""
Show-Separator
Write-Host "File Delete" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "delete_list.csv"
$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load delete_list.csv")
}
if ($items.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Expand Windows-style environment variables (%USERPROFILE%, %TEMP%, etc.)
foreach ($item in $items) {
    $item.TargetPath = Expand-UserEnvironmentVariables $item.TargetPath
}

# Destructive path guard (CLAUDE.md section 8): evaluate once per row;
# consumed by the display, execution and verification steps. Blocked rows
# are recorded as Fail (config error), never deleted - the guard sits
# OUTSIDE the confirmation gate so it also holds under AutoPilot auto-Y.
foreach ($item in $items) {
    $guard = Test-FabriqProtectedPath -Path $item.TargetPath
    $item | Add-Member -NotePropertyName "_GuardBlocked" -NotePropertyValue (-not $guard.IsSafe)
    $item | Add-Member -NotePropertyName "_GuardReason"  -NotePropertyValue $guard.Reason
}

# ========================================
# Step 2: Display deletion targets with status
# ========================================
Show-Info "Deletion targets: $($items.Count) items"
Write-Host ""

$index = 0
foreach ($item in $items) {
    $index++
    $targetPath = $item.TargetPath
    $ifNotFound = if ($item.IfNotFound) { $item.IfNotFound } else { "Skip" }

    if ($item._GuardBlocked) {
        $marker = "[BLOCKED]"
        $markerColor = "Red"
    }
    elseif (Test-Path $targetPath) {
        $isDir = (Get-Item $targetPath -ErrorAction SilentlyContinue) -is [System.IO.DirectoryInfo]
        $typeLabel = if ($isDir) { "Dir" } else { "File" }
        $marker = "[$typeLabel][Exists]"
        $markerColor = "White"
    }
    else {
        $marker = "[Not Found]"
        $markerColor = if ($ifNotFound -eq "Error") { "Red" } else { "Gray" }
    }

    Write-Host "  [$index] $($item.Description)  $marker" -ForegroundColor $markerColor
    Write-Host "      Path: $targetPath" -ForegroundColor DarkGray
    if ($item._GuardBlocked) {
        Write-Host "      Reason: $($item._GuardReason) - will be recorded as Fail" -ForegroundColor Red
    }
    if ($ifNotFound -eq "Error") {
        Write-Host "      IfNotFound: Error" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# ========================================
# Step 3: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Delete the above targets?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 4: Execute deletion
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
$total        = $items.Count
$current      = 0

foreach ($item in $items) {
    $current++
    $targetPath = $item.TargetPath
    $ifNotFound = if ($item.IfNotFound) { $item.IfNotFound } else { "Skip" }

    Write-Host "[$current/$total] $($item.Description)" -ForegroundColor Cyan
    Write-Host "  Path: $targetPath" -ForegroundColor DarkGray

    # Destructive path guard (CLAUDE.md section 8): blocked targets are
    # config errors - record Fail, never delete.
    if ($item._GuardBlocked) {
        Show-Error "Blocked protected path: $targetPath ($($item._GuardReason))"
        $failCount++
        Write-Host ""
        continue
    }

    if (-not (Test-Path $targetPath)) {
        if ($ifNotFound -eq "Error") {
            Show-Error "Target not found: $targetPath"
            $failCount++
        }
        else {
            Show-Skip "Not found — skipped"
            $skipCount++
        }
        Write-Host ""
        continue
    }

    try {
        Remove-Item -Path $targetPath -Force -Recurse -ErrorAction Stop
        Show-Success "Deleted"
        $successCount++
    }
    catch {
        Show-Error "Deletion failed: $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying deleted targets..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($item in $items) {
    $targetPath = $item.TargetPath
    $ifNotFound = if ($item.IfNotFound) { $item.IfNotFound } else { "Skip" }

    # Blocked rows were never deleted (already counted as Fail above);
    # checking "still exists" here would double-penalize them.
    if ($item._GuardBlocked) {
        Write-Host "  [SKIPPED] $($item.Description) (blocked - not deleted)" -ForegroundColor DarkGray
        continue
    }

    if (-not (Test-Path $targetPath)) {
        Write-Host "  [VERIFIED] $($item.Description) (not found)" -ForegroundColor Green
        $verifyPass++
    } else {
        Write-Host "  [VERIFY FAILED] $($item.Description) (still exists)" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)

# ========================================
# Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "File Delete Results" -Verified $verified)
