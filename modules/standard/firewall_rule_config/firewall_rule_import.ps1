# ========================================
# Firewall Rule Import Script
# ========================================
# [PURPOSE]
# Restore a Windows Firewall policy snapshot (.wfw) produced by
# firewall_rule_export.ps1 or by `netsh advfirewall export`. Destructive:
# the operation REPLACES all existing rules, profile states, and IPsec
# settings.
#
# [SAFETY]
# Only rows with IAcknowledgeReplace=1 are executed. Any other value
# rejects the row and counts as a failure (data validation error).
# AutoPilot mode also honors this gate.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Firewall Rule Import" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper: Capture netsh output with correct encoding
# ========================================
# Use the .NET Process API with StandardOutputEncoding=UTF-8. Empirical
# observation on Windows 10/11 (verified on a JP host with consoleCP=932):
# when netsh's stdout is redirected to a pipe, netsh writes its output in
# UTF-8 regardless of the parent's chcp state.
#
# The PS 5.1 `& exe 2>&1` capture path attaches a real console and reads
# via [Console]::OutputEncoding, which can drift out of sync with the
# actual console CP and produce mojibake (UTF-8 bytes decoded as CP932 ->
# the "rule name" kanji becomes unreadable mojibake). Process API + explicit StandardOutput-
# Encoding=UTF8 bypasses that mismatch entirely without mutating any host
# console state.
function Invoke-NetshCapture {
    param([Parameter(Mandatory)][string[]]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'netsh.exe'
    $psi.Arguments = ($Arguments | ForEach-Object {
        if ($_ -match '\s') { '"' + $_ + '"' } else { $_ }
    }) -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

    $p = [System.Diagnostics.Process]::Start($psi)
    $outText = $p.StandardOutput.ReadToEnd()
    $errText = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    $global:LASTEXITCODE = $p.ExitCode

    $combined = $outText
    if (-not [string]::IsNullOrWhiteSpace($errText)) {
        $combined += "`n[stderr]`n$errText"
    }
    return $combined
}

# ========================================
# Helper: Resolve SourcePath
# ========================================
# Resolution rules:
#   1. Absolute path (drive letter / UNC) -> used as-is
#   2. Relative path -> resolved against <module>\backup\
#   3. If the resolved path points to a directory, append policy.wfw
# Operators must NOT include the leading "backup/" segment in relative
# paths, otherwise the result becomes <module>\backup\backup\... and fails.
function Resolve-ImportSource {
    param([string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        $resolved = $Path
    }
    else {
        $resolved = Join-Path (Join-Path $PSScriptRoot "backup") $Path
    }

    if (Test-Path $resolved -PathType Container) {
        $resolved = Join-Path $resolved "policy.wfw"
    }

    return $resolved
}

# ========================================
# Step 0: Privilege Check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required for netsh advfirewall import."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "firewall_rule_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Mode", "SourcePath", "DestinationPath", "IAcknowledgeReplace")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load firewall_rule_list.csv")
}

$importItems = @($enabledItems | Where-Object { $_.Mode -eq "Import" })
if ($importItems.Count -eq 0) {
    Show-Skip "No enabled Import entries"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled Import entries")
}

# ========================================
# Step 2: Acknowledgement Gate + Source Validation
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

$plans = @()
foreach ($item in $importItems) {
    $desc = if ($item.Description) { $item.Description } else { "Firewall import" }

    if ($item.IAcknowledgeReplace -ne "1") {
        Show-Error "[REJECTED] $desc : IAcknowledgeReplace must be '1' for destructive import (got '$($item.IAcknowledgeReplace)')"
        $failCount++
        continue
    }

    if ([string]::IsNullOrWhiteSpace($item.SourcePath)) {
        Show-Error "[REJECTED] $desc : SourcePath is empty"
        $failCount++
        continue
    }

    $resolvedSource = Resolve-ImportSource -Path $item.SourcePath
    if (-not (Test-Path $resolvedSource -PathType Leaf)) {
        Show-Error "[REJECTED] $desc : Source file not found"
        Write-Host "    Original  : $($item.SourcePath)" -ForegroundColor DarkGray
        Write-Host "    Resolved  : $resolvedSource" -ForegroundColor DarkGray
        $failCount++
        continue
    }

    # Try to read sidecar manifest if it sits next to the .wfw file
    $sourceDir   = Split-Path -Parent $resolvedSource
    $manifestPath = Join-Path $sourceDir "manifest.txt"
    $expectedRules = $null
    $sourceOs      = $null
    $manifestExists = Test-Path $manifestPath
    if ($manifestExists) {
        try {
            $manifestText = Get-Content -Path $manifestPath -Raw -ErrorAction Stop
            if ($manifestText -match 'Rule Count\s*:\s*(\d+)') {
                $expectedRules = [int]$Matches[1]
            }
            if ($manifestText -match 'OS Caption\s*:\s*(.+)') {
                $sourceOs = $Matches[1].Trim()
            }
        }
        catch {
            Show-Warning "Could not read manifest at $manifestPath : $_"
            $manifestExists = $false
        }
    }

    # Try to load rule_names.txt sidecar (Name-set verification source)
    $namesFile = Join-Path $sourceDir "rule_names.txt"
    $expectedNames = $null
    if (Test-Path $namesFile) {
        try {
            $expectedNames = @(Get-Content -Path $namesFile -Encoding UTF8 -ErrorAction Stop |
                Where-Object { $_.Trim() -ne '' })
        }
        catch {
            Show-Warning "Could not read rule_names.txt at $namesFile : $_"
        }
    }

    $plans += [PSCustomObject]@{
        Description    = $desc
        OriginalPath   = $item.SourcePath
        ResolvedPath   = $resolvedSource
        ManifestExists = $manifestExists
        ExpectedRules  = $expectedRules
        ExpectedNames  = $expectedNames
        SourceOs       = $sourceOs
    }
}

if ($plans.Count -eq 0) {
    if ($failCount -gt 0) {
        Show-Error "All Import entries were rejected ($failCount item(s))."
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "All Import entries were rejected ($failCount)")
    }
    Show-Skip "No valid Import entries"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid Import entries")
}

# ========================================
# Step 3: Pre-Execution Preview
# ========================================
$currentRuleCount = -1
try {
    $currentRuleCount = @(Get-NetFirewallRule -ErrorAction Stop).Count
}
catch {
    Show-Warning "Could not enumerate current firewall rules: $_"
}

$currentOs = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Planned Imports (DESTRUCTIVE)" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Hostname           : $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  Current rule count : $currentRuleCount" -ForegroundColor White
Write-Host "  Current OS         : $currentOs" -ForegroundColor White
Write-Host ""

foreach ($plan in $plans) {
    Write-Host "  [APPLY] $($plan.Description)" -ForegroundColor Yellow
    Write-Host "    Source         : $($plan.OriginalPath)" -ForegroundColor DarkGray
    if ($plan.OriginalPath -ne $plan.ResolvedPath) {
        Write-Host "      (resolved)   : $($plan.ResolvedPath)" -ForegroundColor DarkGray
    }
    if ($plan.ManifestExists) {
        $rulesText = if ($null -ne $plan.ExpectedRules) { $plan.ExpectedRules } else { "<unknown>" }
        Write-Host "    Expected rules : $rulesText" -ForegroundColor DarkGray
        if ($plan.SourceOs) {
            Write-Host "    Source OS      : $($plan.SourceOs)" -ForegroundColor DarkGray
            if ($currentOs -and $plan.SourceOs -ne $currentOs) {
                Show-Warning "    OS mismatch (source vs current). Import may partially fail."
            }
        }
    }
    else {
        Write-Host "    Manifest       : <not found> (verification will be weak)" -ForegroundColor DarkGray
    }
    if ($null -ne $plan.ExpectedNames) {
        Write-Host "    Verify method  : Name-set ($($plan.ExpectedNames.Count) rule names loaded)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "    Verify method  : count-only (rule_names.txt not present)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Show-Warning "ALL EXISTING RULES, PROFILE STATES, AND IPSEC SETTINGS WILL BE REPLACED."
Write-Host ""

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Replace the firewall policy with the above source(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Apply
# ========================================
$verifyExpectations = @()

foreach ($plan in $plans) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Importing: $($plan.Description)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    $beforeCount = -1
    try {
        $beforeCount = @(Get-NetFirewallRule -ErrorAction Stop).Count
    }
    catch { }

    try {
        $netshOutput = Invoke-NetshCapture -Arguments @('advfirewall', 'import', $plan.ResolvedPath)
        if ($LASTEXITCODE -ne 0) {
            throw "netsh advfirewall import failed (exit=$LASTEXITCODE): $netshOutput"
        }

        $afterCount = -1
        $afterNames = $null
        try {
            $afterRules = @(Get-NetFirewallRule -ErrorAction Stop)
            $afterCount = $afterRules.Count
            $afterNames = @($afterRules | ForEach-Object { $_.Name })
        }
        catch { }

        Show-Success "Imported: $($plan.ResolvedPath)"
        Write-Host "  Before rule count : $beforeCount" -ForegroundColor DarkGray
        Write-Host "  After rule count  : $afterCount" -ForegroundColor DarkGray
        $successCount++

        $verifyExpectations += [PSCustomObject]@{
            Description   = $plan.Description
            BeforeCount   = $beforeCount
            AfterCount    = $afterCount
            ExpectedRules = $plan.ExpectedRules
            ExpectedNames = $plan.ExpectedNames
            AfterNames    = $afterNames
        }
    }
    catch {
        Show-Error "Import failed for '$($plan.Description)': $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
$verified = $null
if ($verifyExpectations.Count -gt 0) {
    Show-Info "Verifying imported policies..."
    Write-Host ""

    $verifyPass = 0
    $verifyFail = 0

    foreach ($exp in $verifyExpectations) {
        $reasons = @()
        $verifyMethod = '<unknown>'

        if ($null -ne $exp.ExpectedNames -and $null -ne $exp.AfterNames) {
            # Preferred: Name-set containment. After-set must be a superset of
            # expected (every snapshotted Name must exist post-import). Extra
            # rules added by Windows' background services after import are
            # ignored - they are not import failures, just dynamic activity.
            $afterSet = New-Object 'System.Collections.Generic.HashSet[string]' (
                [string[]]$exp.AfterNames,
                [System.StringComparer]::OrdinalIgnoreCase
            )
            $missing = @($exp.ExpectedNames | Where-Object { -not $afterSet.Contains($_) })

            if ($missing.Count -eq 0) {
                $verifyMethod = "name-set ($($exp.ExpectedNames.Count) expected, $($exp.AfterNames.Count) actual; superset OK)"
            }
            else {
                $reasons += "$($missing.Count) of $($exp.ExpectedNames.Count) expected rules missing after import"
                if ($missing.Count -le 5) {
                    $reasons += "missing names: $($missing -join ', ')"
                }
                else {
                    $reasons += "missing names (first 5): $(($missing | Select-Object -First 5) -join ', ')"
                }
            }
        }
        elseif ($exp.AfterCount -lt 0) {
            $reasons += "could not read post-import rule count"
        }
        elseif ($null -ne $exp.ExpectedRules) {
            # Fallback: count comparison (for legacy backups without rule_names.txt)
            $verifyMethod = "count-only (legacy backup, no rule_names.txt)"
            if ($exp.AfterCount -ne $exp.ExpectedRules) {
                $reasons += "rule count mismatch (expected=$($exp.ExpectedRules), got=$($exp.AfterCount))"
            }
        }
        else {
            # No manifest, no names: weakest check
            $verifyMethod = "weak (no manifest)"
            if ($exp.AfterCount -le 0) {
                $reasons += "no rules present after import"
            }
        }

        if ($reasons.Count -eq 0) {
            Write-Host "  [VERIFIED] $($exp.Description) [$verifyMethod]" -ForegroundColor Green
            $verifyPass++
        }
        else {
            Write-Host "  [VERIFY FAILED] $($exp.Description) [$verifyMethod] : $($reasons -join '; ')" -ForegroundColor Red
            $verifyFail++
        }
    }

    Write-Host ""
    $verified = ($verifyFail -eq 0)
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Firewall Rule Import Results" -Verified $verified)
