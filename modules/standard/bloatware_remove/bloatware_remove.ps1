# ========================================
# Bloatware Remove Script
# ========================================
# Reads bloatware_list.csv and uninstalls the
# enabled legacy desktop applications by dynamically
# looking up UninstallString from the registry at
# runtime using DisplayName / MatchPattern.
#
# This approach eliminates version-specific path
# mismatches and CSV comma corruption issues that
# occur with static UninstallString storage.
#
# NOTES:
# - Administrator privileges are required.
# - Entries matching NoRemove=1 or SystemComponent=1
#   in the registry are automatically skipped.
# - MatchPattern supports wildcards (-like operator).
#   If omitted, DisplayName is used for exact match.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Bloatware Remove" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Helper: Auto-quote unquoted executable paths
# Some registry UninstallStrings omit quotes around
# paths that contain spaces. cmd.exe /c splits on
# the first space and fails to locate the binary.
# Already-quoted strings are returned unchanged.
# ========================================
function Invoke-QuoteUninstallPath {
    param([string]$CmdString)

    # Already quoted -> return as-is (covers most entries)
    if ($CmdString -match '^"') { return $CmdString }

    # Drive-rooted path: extract up to the .exe/.bat/.cmd/.msi boundary,
    # then quote the executable portion, preserving any trailing arguments.
    if ($CmdString -match '^([A-Za-z]:\\.+?\.(?:exe|bat|cmd|msi))(\s+.*)?$') {
        $exePart = $Matches[1].Trim()
        $argPart = if ($Matches[2]) { $Matches[2].Trim() } else { "" }
        if ($argPart) {
            return "`"$exePart`" $argPart"
        } else {
            return "`"$exePart`""
        }
    }

    # Not a recognized pattern (e.g. bare command name) -> return unchanged
    return $CmdString
}


# ========================================
# Helper: Find matching apps in registry
# Scans HKLM 64-bit and 32-bit Uninstall hives
# for apps whose DisplayName matches the given
# pattern. Follows bloatware_export.ps1 enumeration
# pattern with DisplayName deduplication.
# ========================================
function Find-RegistryUninstallEntry {
    param(
        [Parameter(Mandatory)]
        [string]$Pattern,
        [switch]$ExactMatch
    )

    $regPaths = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*";             Arch = "64bit" }
        @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"; Arch = "32bit" }
    )

    $results   = @()
    $seenNames = @{}

    foreach ($rp in $regPaths) {
        $entries = @(Get-ItemProperty $rp.Path -ErrorAction SilentlyContinue)
        foreach ($entry in $entries) {
            if ([string]::IsNullOrWhiteSpace($entry.DisplayName)) { continue }
            if ($seenNames.ContainsKey($entry.DisplayName))       { continue }

            $matched = if ($ExactMatch) {
                $entry.DisplayName -eq $Pattern
            } else {
                $entry.DisplayName -like $Pattern
            }

            if (-not $matched) { continue }
            $seenNames[$entry.DisplayName] = $true

            $regKey = $entry.PSPath `
                -replace 'Microsoft\.PowerShell\.Core\\Registry::HKEY_LOCAL_MACHINE', 'HKLM:' `
                -replace 'Microsoft\.PowerShell\.Core\\Registry::HKEY_CURRENT_USER',  'HKCU:'

            $results += [PSCustomObject]@{
                DisplayName          = $entry.DisplayName
                Publisher            = if ($entry.Publisher)            { $entry.Publisher }            else { "" }
                DisplayVersion       = if ($entry.DisplayVersion)       { $entry.DisplayVersion }       else { "" }
                Architecture         = $rp.Arch
                WindowsInstaller     = if ($entry.WindowsInstaller)     { "$($entry.WindowsInstaller)" } else { "0" }
                QuietUninstallString = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { "" }
                UninstallString      = if ($entry.UninstallString)      { $entry.UninstallString }      else { "" }
                NoRemove             = if ($entry.NoRemove)             { "$($entry.NoRemove)" }        else { "0" }
                SystemComponent      = if ($entry.SystemComponent)      { "$($entry.SystemComponent)" } else { "0" }
                RegistryKey          = $regKey
            }
        }
    }

    return $results
}


# ========================================
# Step 1: CSV load
# ========================================
$csvPath = Join-Path $PSScriptRoot "bloatware_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "DisplayName")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load bloatware_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Pre-flight check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required to uninstall applications."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}


# ========================================
# Step 3: Pre-execution display
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Uninstall Targets" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($item in $enabledItems) {
    $displayLabel = if ($item.Description) { $item.Description } else { $item.DisplayName }
    $pattern      = if (-not [string]::IsNullOrWhiteSpace($item.MatchPattern)) { $item.MatchPattern } else { $null }
    $useExact     = ($null -eq $pattern)
    $searchKey    = if ($useExact) { $item.DisplayName } else { $pattern }

    $regMatches = @(Find-RegistryUninstallEntry -Pattern $searchKey -ExactMatch:$useExact)

    if ($regMatches.Count -eq 0) {
        Write-Host "  [NOT FOUND] $displayLabel" -ForegroundColor DarkGray
        Write-Host "    (Not installed or already removed)" -ForegroundColor DarkGray
        Write-Host ""
        continue
    }

    foreach ($reg in $regMatches) {
        # Safety skip flags from live registry
        if ($reg.NoRemove -eq "1" -or $reg.SystemComponent -eq "1") {
            $reason = if ($reg.NoRemove -eq "1") { "NoRemove=1" } else { "SystemComponent=1" }
            Write-Host "  [SKIP] $($reg.DisplayName)" -ForegroundColor DarkGray
            Write-Host "    Reason: $reason" -ForegroundColor DarkGray
            Write-Host ""
            continue
        }

        # Determine display method label
        $method = if (-not [string]::IsNullOrWhiteSpace($reg.QuietUninstallString)) {
            "QuietUninstall"
        } elseif ($reg.WindowsInstaller -eq "1") {
            "MSI (/qn)"
        } elseif (-not [string]::IsNullOrWhiteSpace($reg.UninstallString)) {
            "Uninstall"
        } else {
            "No uninstall string"
        }

        Write-Host "  [INSTALLED] $($reg.DisplayName)" -ForegroundColor Yellow
        Write-Host "    Publisher: $($reg.Publisher)" -ForegroundColor DarkGray
        Write-Host "    Version:   $($reg.DisplayVersion)" -ForegroundColor DarkGray
        Write-Host "    Method:    $method" -ForegroundColor DarkGray
        Write-Host ""
    }
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Uninstall the applications listed above?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Uninstall loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    $displayLabel = if ($item.Description) { $item.Description } else { $item.DisplayName }
    $pattern      = if (-not [string]::IsNullOrWhiteSpace($item.MatchPattern)) { $item.MatchPattern } else { $null }
    $useExact     = ($null -eq $pattern)
    $searchKey    = if ($useExact) { $item.DisplayName } else { $pattern }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayLabel" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Re-scan registry (do not cache; prior uninstall may have removed entries)
    $regMatches = @(Find-RegistryUninstallEntry -Pattern $searchKey -ExactMatch:$useExact)

    if ($regMatches.Count -eq 0) {
        Show-Skip "Not found in registry (already uninstalled or not present)"
        Write-Host ""
        $skipCount++
        continue
    }

    foreach ($reg in $regMatches) {
        # ----------------------------------------
        # Safety skip checks (from live registry)
        # ----------------------------------------
        if ($reg.NoRemove -eq "1") {
            Show-Skip "$($reg.DisplayName): NoRemove=1"
            Write-Host ""
            $skipCount++
            continue
        }
        if ($reg.SystemComponent -eq "1") {
            Show-Skip "$($reg.DisplayName): SystemComponent=1"
            Write-Host ""
            $skipCount++
            continue
        }

        # Idempotency: verify registry key still exists
        if (-not (Test-Path $reg.RegistryKey)) {
            Show-Skip "$($reg.DisplayName): Already uninstalled (registry key not found)"
            Write-Host ""
            $skipCount++
            continue
        }

        # ----------------------------------------
        # Determine uninstall method from live registry data
        # ----------------------------------------
        $useMethod = ""
        $cmdString = ""
        $msiGuid   = ""

        if (-not [string]::IsNullOrWhiteSpace($reg.QuietUninstallString)) {
            # Priority 1: QuietUninstallString
            $useMethod = "quiet"
            $cmdString = $reg.QuietUninstallString
        } elseif ($reg.WindowsInstaller -eq "1" -and -not [string]::IsNullOrWhiteSpace($reg.UninstallString)) {
            # Priority 2: MSI — extract GUID for silent uninstall
            if ($reg.UninstallString -match '\{[0-9A-Fa-f\-]+\}') {
                $useMethod = "msi"
                $msiGuid   = $Matches[0]
            } else {
                Show-Warning "MSI GUID not found in UninstallString. Falling back to standard method."
                $useMethod = "standard"
                $cmdString = $reg.UninstallString
            }
        } elseif (-not [string]::IsNullOrWhiteSpace($reg.UninstallString)) {
            # Priority 3: UninstallString as-is
            $useMethod = "standard"
            $cmdString = $reg.UninstallString
        } else {
            Show-Skip "$($reg.DisplayName): No uninstall string available"
            Write-Host ""
            $skipCount++
            continue
        }

        # ----------------------------------------
        # Execute uninstall
        # ----------------------------------------
        try {
            $proc = $null

            switch ($useMethod) {
                "quiet" {
                    $cmdToRun = Invoke-QuoteUninstallPath $cmdString
                    Show-Info "Method: QuietUninstall"
                    Show-Info "Command: $cmdToRun"
                    $proc = Start-Process -FilePath "cmd.exe" `
                        -ArgumentList @("/c", $cmdToRun) `
                        -Wait -NoNewWindow -PassThru -ErrorAction Stop
                }
                "msi" {
                    $msiArgs = "/X$msiGuid /qn /norestart"
                    Show-Info "Method: MSI (/qn)"
                    Show-Info "Command: msiexec.exe $msiArgs"
                    $proc = Start-Process -FilePath "msiexec.exe" `
                        -ArgumentList $msiArgs `
                        -Wait -NoNewWindow -PassThru -ErrorAction Stop
                }
                "standard" {
                    $cmdToRun = Invoke-QuoteUninstallPath $cmdString
                    Show-Info "Method: Uninstall"
                    Show-Info "Command: $cmdToRun"
                    $proc = Start-Process -FilePath "cmd.exe" `
                        -ArgumentList @("/c", $cmdToRun) `
                        -Wait -NoNewWindow -PassThru -ErrorAction Stop
                }
            }

            switch ($proc.ExitCode) {
                0 {
                    Show-Success "$($reg.DisplayName) uninstalled successfully (ExitCode: 0)"
                    $successCount++
                }
                3010 {
                    Show-Success "$($reg.DisplayName) uninstalled successfully (ExitCode: 3010)"
                    Show-Warning "A reboot is required to complete the uninstallation."
                    $successCount++
                }
                default {
                    Show-Error "$($reg.DisplayName) completed with ExitCode: $($proc.ExitCode)"
                    Show-Info "Check application logs for details."
                    $failCount++
                }
            }
        }
        catch {
            Show-Error "Failed: $($reg.DisplayName) : $_"
            $failCount++
        }

        Write-Host ""
    }
}


# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Show-Info "Verifying uninstalled apps..."
Write-Host ""

$verifyPass = 0
$verifyFail = 0

foreach ($item in $enabledItems) {
    $displayLabel = if ($item.Description) { $item.Description } else { $item.DisplayName }
    $pattern      = if (-not [string]::IsNullOrWhiteSpace($item.MatchPattern)) { $item.MatchPattern } else { $null }
    $useExact     = ($null -eq $pattern)
    $searchKey    = if ($useExact) { $item.DisplayName } else { $pattern }

    $remaining = @(Find-RegistryUninstallEntry -Pattern $searchKey -ExactMatch:$useExact |
        Where-Object { $_.NoRemove -ne "1" -and $_.SystemComponent -ne "1" })

    if ($remaining.Count -eq 0) {
        Write-Host "  [VERIFIED] $displayLabel (not found in registry)" -ForegroundColor Green
        $verifyPass++
    } else {
        $names = ($remaining | ForEach-Object { $_.DisplayName }) -join ", "
        Write-Host "  [VERIFY FAILED] $displayLabel (still installed: $names)" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)


# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Bloatware Remove Results" -Verified $verified)
