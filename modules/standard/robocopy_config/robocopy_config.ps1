# ========================================
# Robocopy Config Script
# ========================================
# Executes Robocopy jobs defined in robocopy_list.csv.
# Arguments are built from CSV flag columns (Recursive, Mirror,
# CopyACL, SkipOlder) with a CustomOptions escape hatch.
#
# [NOTES]
# - Requires administrator privileges for some operations
# - Robocopy ExitCode 0-7 = success/warning, 8+ = failure
# - Source must exist; Destination reachability is delegated to Robocopy
# - Baseline /R:3 /W:5 /NP is always applied (safeguard)
# - UNC authentication via optional AuthUser/AuthPass columns
# ========================================

Write-Host ""
Show-Separator
Write-Host "Robocopy" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "robocopy_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "ID", "Source", "Destination", "Recursive", "Mirror", "CopyACL", "SkipOlder")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load robocopy_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Pre-flight check
# ========================================
$robocopyPath = Get-Command "robocopy.exe" -ErrorAction SilentlyContinue
if (-not $robocopyPath) {
    Show-Error "robocopy.exe not found on this system."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "robocopy.exe not found")
}


# ========================================
# Step 3: Pre-execution display
# ========================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Robocopy Jobs" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { "Job $($item.ID)" }

    # Source existence check for display
    $srcExists = Test-Path $item.Source

    if ($srcExists) {
        Write-Host "  [READY] $displayName" -ForegroundColor Yellow
    }
    else {
        Write-Host "  [SRC NOT FOUND] $displayName" -ForegroundColor Red
    }

    Write-Host "    ID:   $($item.ID)" -ForegroundColor DarkGray
    Write-Host "    Src:  $($item.Source)" -ForegroundColor DarkGray
    Write-Host "    Dst:  $($item.Destination)" -ForegroundColor DarkGray

    # Build flag summary for display
    $flags = @()
    if ($item.Mirror -eq "1")    { $flags += "MIR" }
    elseif ($item.Recursive -eq "1") { $flags += "E" }
    if ($item.CopyACL -eq "1")   { $flags += "COPYALL" }
    if ($item.SkipOlder -eq "1") { $flags += "XO" }
    $flagStr = if ($flags.Count -gt 0) { ($flags -join ", ") } else { "(none)" }
    Write-Host "    Flags: $flagStr" -ForegroundColor DarkGray

    if (-not [string]::IsNullOrWhiteSpace($item.CustomOptions)) {
        Write-Host "    Custom: $($item.CustomOptions)" -ForegroundColor DarkGray
    }

    # Show auth info (user only, never password)
    if (-not [string]::IsNullOrWhiteSpace($item.AuthUser)) {
        Write-Host "    Auth: $($item.AuthUser)" -ForegroundColor DarkGray
    }

    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Execute the above Robocopy jobs?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execution loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayName = if ($item.Description) { $item.Description } else { "Job $($item.ID)" }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Executing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # ----------------------------------------
    # Source existence check (outside try)
    # ----------------------------------------
    if (-not (Test-Path $item.Source)) {
        Show-Error "Source not found: $($item.Source)"
        Write-Host ""
        $failCount++
        continue
    }

    # ----------------------------------------
    # UNC authentication: establish net use sessions
    # ----------------------------------------
    $connectedShares = @()
    $authFailed = $false
    $useAuth = -not [string]::IsNullOrWhiteSpace($item.AuthUser) -and
               -not [string]::IsNullOrWhiteSpace($item.AuthPass)

    if ($useAuth) {
        # Extract unique UNC share roots from Source and Destination
        $uncShares = @()
        foreach ($path in @($item.Source, $item.Destination)) {
            if ("$path" -match '^(\\\\[^\\]+\\[^\\]+)') {
                $share = $Matches[1]
                if ($uncShares -inotcontains $share) {
                    $uncShares += $share
                }
            }
        }

        # Establish net use connections
        foreach ($share in $uncShares) {
            Show-Info "Authenticating: $share (User: $($item.AuthUser))"

            $netOutput = & net use $share "$($item.AuthPass)" /user:"$($item.AuthUser)" 2>&1
            $netExitCode = $LASTEXITCODE

            if ($netExitCode -ne 0) {
                Show-Error "net use failed (ExitCode=$netExitCode): $share"
                foreach ($line in $netOutput) {
                    Write-Host "  $line" -ForegroundColor DarkGray
                }
                $authFailed = $true
                break
            }

            Show-Success "Connected: $share"
            $connectedShares += $share
        }
    }

    # If authentication failed, cleanup and skip this job
    if ($authFailed) {
        foreach ($s in $connectedShares) {
            & net use $s /delete /y 2>&1 | Out-Null
        }
        Write-Host ""
        $failCount++
        continue
    }

    # ----------------------------------------
    # Main processing: Robocopy execution
    # ----------------------------------------
    try {
        Show-Info "Source:      $($item.Source)"
        Show-Info "Destination: $($item.Destination)"

        # Build arguments from CSV flags
        # Baseline: safeguard against default retry hell + progress suppression
        $arguments = "`"$($item.Source)`" `"$($item.Destination)`" /R:3 /W:5 /NP"

        # Mirror vs Recursive (Mirror takes priority)
        if ($item.Mirror -eq "1") {
            $arguments += " /MIR"
        }
        elseif ($item.Recursive -eq "1") {
            $arguments += " /E"
        }

        # Copy attributes
        if ($item.CopyACL -eq "1") {
            $arguments += " /COPYALL /DCOPY:DAT"
        }
        else {
            $arguments += " /COPY:DAT /DCOPY:DAT"
        }

        # Skip older files
        if ($item.SkipOlder -eq "1") {
            $arguments += " /XO"
        }

        # Escape hatch: custom options (appended last)
        if (-not [string]::IsNullOrWhiteSpace($item.CustomOptions)) {
            $arguments += " $($item.CustomOptions)"
        }

        # Display final command line for verification
        Show-Info "Arguments:  $arguments"

        $proc = Start-Process -FilePath "robocopy.exe" `
            -ArgumentList $arguments `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop

        $exitCode = $proc.ExitCode

        # ----------------------------------------
        # Exit code evaluation
        # Robocopy: 0-7 = success/warning, 8+ = failure
        # ----------------------------------------
        if ($exitCode -le 1) {
            # 0: No change, 1: Files copied
            Show-Success "Completed (ExitCode=$exitCode): $displayName"
            $successCount++
        }
        elseif ($exitCode -le 3) {
            # 2: Extra files, 3: Copied + Extra
            Show-Success "Completed with extra files (ExitCode=$exitCode): $displayName"
            $successCount++
        }
        elseif ($exitCode -le 7) {
            # 4-7: Mismatched files detected
            Show-Warning "Completed with mismatches (ExitCode=$exitCode): $displayName"
            $successCount++
        }
        else {
            # 8+: Serious error
            Show-Error "Failed (ExitCode=$exitCode): $displayName"
            $failCount++
        }
    }
    catch {
        Show-Error "Failed: $displayName : $_"
        $failCount++
    }
    finally {
        # ----------------------------------------
        # UNC cleanup: disconnect all shares (silent)
        # ----------------------------------------
        foreach ($share in $connectedShares) {
            & net use $share /delete /y 2>&1 | Out-Null
        }
    }

    Write-Host ""
}


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Robocopy Results")
