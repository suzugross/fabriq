# ========================================
# Firewall Rule Export Script
# ========================================
# [PURPOSE]
# Snapshot the current Windows Firewall policy (rules + profile state +
# logging + IPsec) to a portable .wfw file plus human-readable sidecar
# manifest files. Non-destructive.
#
# [OUTPUTS] (per snapshot directory)
#   policy.wfw        - netsh advfirewall export (restore source of truth)
#   rules_show.txt    - netsh advfirewall firewall show rule name=all verbose
#   rules.json        - Get-NetFirewallRule subset, JSON-formatted
#   profiles.json     - Get-NetFirewallProfile, JSON-formatted
#   manifest.txt      - Hostname / OS / Rule count / Profile states / Hash
# ========================================

Write-Host ""
Show-Separator
Write-Host "Firewall Rule Export" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper: Capture netsh output with correct encoding
# ========================================
# Use the .NET Process API with StandardOutputEncoding=UTF-8. Empirical
# observation on Windows 10/11 (verified on a JP host with consoleCP=932):
# when netsh's stdout is redirected to a pipe, netsh writes its output in
# UTF-8 regardless of the parent's chcp state - querying the raw bytes
# from the redirected pipe shows UTF-8 sequences that correctly decode to
# Japanese rule names (e.g. bytes E8 A6 8F E5 89 87 E5 90 8D = the "rule name" kanji).
#
# The PS 5.1 `& exe 2>&1` capture path attaches a real console and reads
# via [Console]::OutputEncoding, which can drift out of sync with the
# actual console CP and produce mojibake (UTF-8 bytes decoded as CP932 ->
# the "rule name" kanji becomes unreadable mojibake). Process API + explicit StandardOutput-
# Encoding=UTF8 bypasses that mismatch entirely without mutating any host
# console state.
#
# Note: this differs from Invoke-CScriptCapture in evidence_config, which
# uses the OEM codepage - cscript's redirected-stdout encoding behavior
# differs from netsh's. The choice here is netsh-specific and validated
# empirically.
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
# Helper: Auto-register Import entry into firewall_rule_list.csv
# ========================================
# Append a single Import row to the CSV pointing at the freshly created
# policy.wfw, with Enabled=0 and IAcknowledgeReplace=0 (gates so it does
# not fire on the next run unless the operator explicitly turns it on).
# Encoding (UTF-8 BOM vs ANSI) and trailing-newline state are detected
# from the existing file and preserved.
function Add-ImportEntry {
    param(
        [Parameter(Mandatory)][string]$CsvPath,
        [Parameter(Mandatory)][string]$WfwAbsolutePath,
        [Parameter(Mandatory)][string]$BackupRoot,
        [string]$Segment = ''
    )

    if (-not (Test-Path $CsvPath)) {
        Show-Warning "Cannot auto-register Import entry: CSV not found at $CsvPath"
        return
    }

    # Compute SourcePath: relative form when wfw lives under <module>\backup\,
    # absolute path otherwise. The Import script's Resolve-ImportSource
    # treats relative as anchored under <module>\backup\.
    $resolvedWfw  = [System.IO.Path]::GetFullPath($WfwAbsolutePath)
    $resolvedRoot = [System.IO.Path]::GetFullPath($BackupRoot).TrimEnd('\') + '\'
    if ($resolvedWfw.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $sourcePath = ($resolvedWfw.Substring($resolvedRoot.Length)) -replace '\\', '/'
    }
    else {
        $sourcePath = $resolvedWfw
    }

    # Read header to determine column order
    try {
        $headerLine = Get-Content -Path $CsvPath -TotalCount 1 -ErrorAction Stop
        $columns = $headerLine -split ','
    }
    catch {
        Show-Warning "Cannot auto-register Import entry: failed to read CSV header: $_"
        return
    }

    # Detect encoding from BOM and trailing-newline state from raw bytes
    $bytes = [System.IO.File]::ReadAllBytes($CsvPath)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $encoding = if ($hasBom) { 'UTF8' } else { 'Default' }
    $endsWithNewline = $bytes.Length -gt 0 -and ($bytes[$bytes.Length - 1] -eq 0x0A -or $bytes[$bytes.Length - 1] -eq 0x0D)

    # Build the row by column name; unknown columns become empty
    $description = "Auto-registered from Export at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $values = @{
        'Enabled'             = '0'
        'Mode'                = 'Import'
        'SourcePath'          = $sourcePath
        'DestinationPath'     = ''
        'IAcknowledgeReplace' = '0'
        'Description'         = $description
        'Segment'             = $Segment
    }

    $fields = foreach ($col in $columns) {
        $v = if ($values.ContainsKey($col)) { $values[$col] } else { '' }
        if ($null -eq $v) { $v = '' }
        if ($v -match '[",\r\n]') {
            '"' + ($v -replace '"', '""') + '"'
        } else {
            $v
        }
    }
    $newLine = $fields -join ','

    try {
        $valueToAppend = if ($endsWithNewline) { $newLine } else { "`r`n$newLine" }
        Add-Content -Path $CsvPath -Value $valueToAppend -Encoding $encoding -ErrorAction Stop
        Show-Info "Auto-registered Import entry: $sourcePath"
    }
    catch {
        Show-Warning "Failed to auto-register Import entry (continuing): $_"
    }
}

# ========================================
# Step 0: Privilege Check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required for netsh advfirewall export."
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

$exportItems = @($enabledItems | Where-Object { $_.Mode -eq "Export" })
if ($exportItems.Count -eq 0) {
    Show-Skip "No enabled Export entries"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled Export entries")
}

# ========================================
# Step 2: Resolve Destinations
# ========================================
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$defaultBackupRoot = Join-Path $PSScriptRoot "backup"

$plans = @()
foreach ($item in $exportItems) {
    $dest = $item.DestinationPath
    if ([string]::IsNullOrWhiteSpace($dest)) {
        $dest = Join-Path $defaultBackupRoot $timestamp
    }
    $plans += [PSCustomObject]@{
        Description = if ($item.Description) { $item.Description } else { "Firewall snapshot" }
        Destination = $dest
        Segment     = if ($null -ne $item.PSObject.Properties['Segment']) { $item.Segment } else { '' }
    }
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

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Planned Snapshots" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Hostname           : $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  Current rule count : $currentRuleCount" -ForegroundColor White
Write-Host ""

foreach ($plan in $plans) {
    Write-Host "  [APPLY] $($plan.Description)" -ForegroundColor Yellow
    Write-Host "    Destination: $($plan.Destination)" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Export the firewall policy to the above destination(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Apply
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0
$verifyExpectations = @()

foreach ($plan in $plans) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Exporting: $($plan.Description)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        if (-not (Test-Path $plan.Destination)) {
            New-Item -Path $plan.Destination -ItemType Directory -Force | Out-Null
        }

        $wfwPath      = Join-Path $plan.Destination "policy.wfw"
        $showPath     = Join-Path $plan.Destination "rules_show.txt"
        $rulesJson    = Join-Path $plan.Destination "rules.json"
        $profilesJson = Join-Path $plan.Destination "profiles.json"
        $manifestPath = Join-Path $plan.Destination "manifest.txt"

        # 1) netsh export -> policy.wfw
        $netshOutput = Invoke-NetshCapture -Arguments @('advfirewall', 'export', $wfwPath)
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $wfwPath)) {
            throw "netsh advfirewall export failed (exit=$LASTEXITCODE): $netshOutput"
        }

        # 2) netsh show rule -> rules_show.txt (audit trail)
        $showOutput = Invoke-NetshCapture -Arguments @('advfirewall', 'firewall', 'show', 'rule', 'name=all', 'verbose')
        Set-Content -Path $showPath -Value $showOutput -Encoding UTF8

        # 3) Get-NetFirewallRule -> rules.json (curated property subset to avoid CIM bloat)
        $rulesArray = @(Get-NetFirewallRule -ErrorAction Stop)
        try {
            $rulesArray |
                Select-Object DisplayName, Name, Description, Group, Enabled, Profile, `
                              Direction, Action, EdgeTraversalPolicy, LooseSourceMapping, `
                              LocalOnlyMapping, Owner, PolicyStoreSource, PolicyStoreSourceType |
                ConvertTo-Json -Depth 5 |
                Set-Content -Path $rulesJson -Encoding UTF8
        }
        catch {
            Show-Warning "rules.json generation failed (continuing): $_"
        }

        # 4) Get-NetFirewallProfile -> profiles.json
        $profilesArray = @(Get-NetFirewallProfile -ErrorAction Stop)
        try {
            $profilesArray | ConvertTo-Json -Depth 5 | Set-Content -Path $profilesJson -Encoding UTF8
        }
        catch {
            Show-Warning "profiles.json generation failed (continuing): $_"
        }

        # 5) rule_names.txt - one Name per line, used by Import for
        #    Name-set containment verification (robust against Windows'
        #    background dynamic rule add/remove activity that makes rule
        #    counts noisy across snapshot/restore boundaries).
        $namesPath = Join-Path $plan.Destination "rule_names.txt"
        try {
            $rulesArray | ForEach-Object { $_.Name } |
                Set-Content -Path $namesPath -Encoding UTF8
        }
        catch {
            Show-Warning "rule_names.txt generation failed (continuing): $_"
        }

        # 6) manifest.txt
        $wfwHash = (Get-FileHash -Path $wfwPath -Algorithm SHA256).Hash
        $wfwSize = (Get-Item $wfwPath).Length
        $osCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
        $osVersion = [System.Environment]::OSVersion.Version.ToString()
        $profileLines = $profilesArray | ForEach-Object {
            "  $($_.Name): Enabled=$($_.Enabled), DefaultInbound=$($_.DefaultInboundAction), DefaultOutbound=$($_.DefaultOutboundAction), LogAllowed=$($_.LogAllowed), LogBlocked=$($_.LogBlocked)"
        }
        $manifest = @(
            "Fabriq Firewall Snapshot Manifest"
            "Timestamp        : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
            "Hostname         : $env:COMPUTERNAME"
            "OS Caption       : $osCaption"
            "OS Version       : $osVersion"
            "Rule Count       : $($rulesArray.Count)"
            "WFW File         : policy.wfw"
            "WFW Size (bytes) : $wfwSize"
            "WFW SHA256       : $wfwHash"
            "Names File       : rule_names.txt"
            "Profile States   :"
        ) + $profileLines
        Set-Content -Path $manifestPath -Value $manifest -Encoding UTF8

        Show-Success "Snapshot saved: $($plan.Destination)"
        Write-Host "  policy.wfw size : $wfwSize bytes" -ForegroundColor DarkGray
        Write-Host "  rule count      : $($rulesArray.Count)" -ForegroundColor DarkGray
        $successCount++

        $verifyExpectations += [PSCustomObject]@{
            Description   = $plan.Description
            WfwPath       = $wfwPath
            ManifestPath  = $manifestPath
            NamesPath     = $namesPath
            ExpectedRules = $rulesArray.Count
        }

        # Auto-register Import entry into firewall_rule_list.csv (Enabled=0)
        Add-ImportEntry -CsvPath $csvPath -WfwAbsolutePath $wfwPath `
            -BackupRoot $defaultBackupRoot -Segment $plan.Segment
    }
    catch {
        Show-Error "Export failed for '$($plan.Description)': $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
$verified = $null
if ($verifyExpectations.Count -gt 0) {
    Show-Info "Verifying exported snapshots..."
    Write-Host ""

    $verifyPass = 0
    $verifyFail = 0
    $minSize = 1024  # heuristic: a healthy .wfw is far larger than 1KB

    # Note: we deliberately do NOT compare a fresh Get-NetFirewallRule count
    # to the recorded rule count here. Windows' background services (mpssvc /
    # AppX / GPO) add and remove rules on their own schedule, so a delta
    # between snapshot time and verify time (a few hundred ms later) is noise,
    # not a real failure. File integrity + Names file line-count consistency
    # is sufficient evidence that the export captured a coherent snapshot.

    foreach ($exp in $verifyExpectations) {
        $reasons = @()

        if (-not (Test-Path $exp.WfwPath)) {
            $reasons += "wfw missing"
        }
        else {
            $size = (Get-Item $exp.WfwPath).Length
            if ($size -lt $minSize) {
                $reasons += "wfw too small ($size bytes)"
            }
        }

        if (-not (Test-Path $exp.ManifestPath)) {
            $reasons += "manifest missing"
        }

        if (-not (Test-Path $exp.NamesPath)) {
            $reasons += "rule_names.txt missing"
        }
        else {
            $namesLines = @(Get-Content -Path $exp.NamesPath -Encoding UTF8 | Where-Object { $_.Trim() -ne '' })
            if ($namesLines.Count -ne $exp.ExpectedRules) {
                $reasons += "rule_names.txt line count ($($namesLines.Count)) != recorded rule count ($($exp.ExpectedRules))"
            }
        }

        if ($reasons.Count -eq 0) {
            Write-Host "  [VERIFIED] $($exp.Description)" -ForegroundColor Green
            $verifyPass++
        }
        else {
            Write-Host "  [VERIFY FAILED] $($exp.Description) : $($reasons -join '; ')" -ForegroundColor Red
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
    -Title "Firewall Rule Export Results" -Verified $verified)
