# ========================================
# Directory Cleaner Script
# ========================================
# Reads clean_list.csv and recursively deletes
# the specified directories or their contents.
#
# NOTES:
# - Administrator privileges are recommended for
#   deleting system-level paths.
# - Hardcoded forbidden path whitelist prevents
#   accidental deletion of critical system paths.
# - TargetPath supports Windows environment variables
#   (e.g. %USERPROFILE%, %LOCALAPPDATA%).
# - Mode "contents" removes items inside the folder.
#   Mode "directory" removes the folder itself.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Directory Cleaner" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Hardcoded Forbidden Path Whitelist
# ========================================
$forbiddenPaths = @(
    "C:\",
    "C:\Windows",
    "C:\Windows\System32",
    "C:\Windows\SysWOW64",
    "C:\Windows\WinSxS",
    "C:\Program Files",
    "C:\Program Files (x86)",
    "C:\ProgramData",
    "C:\Users",
    "C:\Recovery",
    "C:\Boot",
    "$env:SystemRoot",
    "$env:SystemRoot\System32",
    "$env:ProgramFiles",
    "${env:ProgramFiles(x86)}",
    "$env:USERPROFILE",
    "$env:PUBLIC"
)

# Normalize forbidden paths: expand env vars → GetFullPath → lower
$normalizedForbidden = @()
foreach ($fp in $forbiddenPaths) {
    try {
        $expanded = [System.Environment]::ExpandEnvironmentVariables($fp)
        $resolved = [System.IO.Path]::GetFullPath($expanded).TrimEnd('\').ToLowerInvariant()
        $normalizedForbidden += $resolved
    }
    catch { }
}

# Also add the fabriq root itself
try {
    $fabriqRoot = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\..\..").TrimEnd('\').ToLowerInvariant()
    $normalizedForbidden += $fabriqRoot
}
catch { }

$normalizedForbidden = $normalizedForbidden | Select-Object -Unique


# ========================================
# Helper: Test-ForbiddenPath
# Returns $true if the path must not be touched.
# ========================================
function Test-ForbiddenPath {
    param([string]$NormalizedPath)

    # Check A: exact match with a forbidden path
    if ($normalizedForbidden -contains $NormalizedPath) {
        return $true
    }

    # Check B: target IS a parent of a forbidden path (e.g. C:\Win → blocked)
    foreach ($fp in $normalizedForbidden) {
        if ($fp.StartsWith($NormalizedPath + "\")) {
            return $true
        }
    }

    # Check C: minimum depth — require at least 3 path segments
    # e.g. C:\Users\John qualifies (3 segments); C:\Users does not (2 segments)
    $segments = $NormalizedPath.Split('\') | Where-Object { $_ -ne "" }
    if ($segments.Count -lt 3) {
        return $true
    }

    return $false
}


# ========================================
# Step 1: CSV load
# ========================================
$csvPath = Join-Path $PSScriptRoot "clean_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "TargetPath", "Mode")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load clean_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# Expand environment variables in TargetPath
foreach ($item in $enabledItems) {
    $item.TargetPath = [System.Environment]::ExpandEnvironmentVariables($item.TargetPath)
}


# ========================================
# Step 2: Validate Mode values
# ========================================
$invalidItems = $enabledItems | Where-Object { $_.Mode -notin @("contents", "directory") }
if ($invalidItems.Count -gt 0) {
    foreach ($inv in $invalidItems) {
        Show-Error "Invalid Mode '$($inv.Mode)' for entry: $($inv.TargetPath)"
    }
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid Mode value(s) in clean_list.csv — must be 'contents' or 'directory'")
}


# ========================================
# Step 3: Pre-execution display (dry-run)
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Deletion Targets" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetPath }

    # Normalize path for safety check
    $normalizedTarget = ""
    try {
        $normalizedTarget = [System.IO.Path]::GetFullPath($item.TargetPath).TrimEnd('\').ToLowerInvariant()
    }
    catch {
        Write-Host "  [ERROR] $displayName" -ForegroundColor Red
        Write-Host "    Invalid path: $($item.TargetPath)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    # Safety check
    if (Test-ForbiddenPath -NormalizedPath $normalizedTarget) {
        Write-Host "  [BLOCKED] $displayName" -ForegroundColor Red
        Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
        Write-Host "    Reason: Protected/system path — will be skipped" -ForegroundColor Red
        Write-Host ""
        continue
    }

    if (-not (Test-Path $item.TargetPath)) {
        Write-Host "  [NOT FOUND] $displayName" -ForegroundColor DarkGray
        Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
        Write-Host "    (Already absent — will be skipped)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    $modeLabel = switch ($item.Mode) {
        "contents"  { "Delete contents (folder retained)" }
        "directory" { "Delete entire directory" }
    }

    Write-Host "  [DELETE] $displayName" -ForegroundColor Yellow
    Write-Host "    Path: $($item.TargetPath)" -ForegroundColor DarkGray
    Write-Host "    Mode: $modeLabel" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""
Show-Warning "The above paths will be permanently deleted. This cannot be undone."
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with directory deletion?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Deletion loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.TargetPath }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Normalize path
    # ----------------------------------------
    $normalizedTarget = ""
    try {
        $normalizedTarget = [System.IO.Path]::GetFullPath($item.TargetPath).TrimEnd('\').ToLowerInvariant()
    }
    catch {
        Show-Error "Invalid path: $($item.TargetPath)"
        Write-Host ""
        $failCount++
        continue
    }

    # ----------------------------------------
    # Safety check (double-gate at execution time)
    # ----------------------------------------
    if (Test-ForbiddenPath -NormalizedPath $normalizedTarget) {
        Show-Skip "Blocked: protected/system path — $($item.TargetPath)"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # Idempotency: path does not exist
    # ----------------------------------------
    if (-not (Test-Path $item.TargetPath)) {
        Show-Skip "Already absent: $($item.TargetPath)"
        Write-Host ""
        $skipCount++
        continue
    }

    # ----------------------------------------
    # Mode: contents — check if already empty
    # ----------------------------------------
    if ($item.Mode -eq "contents") {
        $children = Get-ChildItem $item.TargetPath -Force -ErrorAction SilentlyContinue
        if ($null -eq $children -or $children.Count -eq 0) {
            Show-Skip "Already empty: $($item.TargetPath)"
            Write-Host ""
            $skipCount++
            continue
        }
    }

    # ----------------------------------------
    # Execute deletion
    # ----------------------------------------
    try {
        switch ($item.Mode) {
            "contents" {
                Show-Info "Deleting contents of: $($item.TargetPath)"
                $errors = 0
                $deleted = 0
                $children = Get-ChildItem $item.TargetPath -Force -ErrorAction SilentlyContinue
                foreach ($child in $children) {
                    try {
                        Remove-Item $child.FullName -Recurse -Force -ErrorAction Stop
                        $deleted++
                    }
                    catch {
                        $errors++
                    }
                }
                if ($errors -eq 0) {
                    Show-Success "Contents deleted ($deleted items)"
                    $successCount++
                }
                else {
                    Show-Warning "Contents partially deleted ($deleted deleted, $errors locked/failed)"
                    $successCount++
                }
            }
            "directory" {
                Show-Info "Deleting directory: $($item.TargetPath)"
                Remove-Item $item.TargetPath -Recurse -Force -ErrorAction Stop
                Show-Success "Directory deleted: $($item.TargetPath)"
                $successCount++
            }
        }
    }
    catch {
        Show-Error "Failed: $($_.Exception.Message)"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Directory Cleaner Results")
