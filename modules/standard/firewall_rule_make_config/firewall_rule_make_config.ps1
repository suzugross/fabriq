# ========================================
# Firewall Rule Maker Script
# ========================================
# [PURPOSE]
# Create individual Windows Firewall rules from CSV definitions using
# New-NetFirewallRule. Idempotent: rules whose DisplayName already exists
# are skipped. Per-row validation rejects malformed entries before any
# cmdlet call.
#
# [CSV SCHEMA]
# See Guide.txt for the full column list (17 columns).
# Required: Enabled, DisplayName, Direction, Action.
# Multi-value cells use ';' as the in-cell separator.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Firewall Rule Maker" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Validation tables
# ========================================
$script:ValidDirections     = @('Inbound', 'Outbound')
$script:ValidActions        = @('Allow', 'Block')
$script:ValidProfiles       = @('Domain', 'Private', 'Public', 'Any')
$script:ValidProtocols      = @('TCP', 'UDP', 'ICMPv4', 'ICMPv6', 'Any')
$script:ValidEdgePolicies   = @('Block', 'Allow', 'DeferToUser', 'DeferToApp')
$script:ValidInterfaceTypes = @('Any', 'Wired', 'Wireless', 'RemoteAccess')

# ========================================
# Helper: Split-MultiValue
# ========================================
# Turn a ';'-separated cell into a clean trimmed string array.
# Empty/whitespace input returns @().
function Split-MultiValue {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    return @(($Value -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# ========================================
# Helper: Test-RowValidity
# ========================================
# Validate a single CSV row. Returns array of issue strings (empty = valid).
function Test-RowValidity {
    param($Row)
    $issues = @()

    if ([string]::IsNullOrWhiteSpace($Row.DisplayName)) {
        $issues += "DisplayName is empty"
    }
    if ($Row.Direction -notin $script:ValidDirections) {
        $issues += "Direction must be Inbound or Outbound (got: '$($Row.Direction)')"
    }
    if ($Row.Action -notin $script:ValidActions) {
        $issues += "Action must be Allow or Block (got: '$($Row.Action)')"
    }

    foreach ($p in (Split-MultiValue -Value $Row.Profile)) {
        if ($p -notin $script:ValidProfiles) {
            $issues += "Profile contains invalid value '$p' (must be Domain/Private/Public/Any)"
        }
    }

    if ($Row.Protocol -and $Row.Protocol -notin $script:ValidProtocols) {
        $protoNum = 0
        if (-not [int]::TryParse($Row.Protocol, [ref]$protoNum) -or $protoNum -lt 0 -or $protoNum -gt 255) {
            $issues += "Protocol must be TCP/UDP/ICMPv4/ICMPv6/Any/0-255 (got: '$($Row.Protocol)')"
        }
    }

    foreach ($portCol in @('LocalPort', 'RemotePort')) {
        foreach ($part in (Split-MultiValue -Value $Row.$portCol)) {
            if ($part -eq 'Any') { continue }
            if ($part -match '^\d+$') {
                $n = [int]$part
                if ($n -lt 0 -or $n -gt 65535) {
                    $issues += "$portCol value out of range: $part"
                }
                continue
            }
            if ($part -match '^(\d+)-(\d+)$') {
                $low = [int]$Matches[1]; $high = [int]$Matches[2]
                if ($low -lt 0 -or $high -gt 65535 -or $low -gt $high) {
                    $issues += "$portCol range invalid: $part"
                }
                continue
            }
            $issues += "$portCol value not parseable as port/range/Any: $part"
        }
    }

    if ($Row.RuleEnabled -and $Row.RuleEnabled -notin @('True', 'False')) {
        $issues += "RuleEnabled must be True or False (got: '$($Row.RuleEnabled)')"
    }

    if ($Row.EdgeTraversalPolicy -and $Row.EdgeTraversalPolicy -notin $script:ValidEdgePolicies) {
        $issues += "EdgeTraversalPolicy must be Block/Allow/DeferToUser/DeferToApp (got: '$($Row.EdgeTraversalPolicy)')"
    }

    foreach ($it in (Split-MultiValue -Value $Row.InterfaceType)) {
        if ($it -notin $script:ValidInterfaceTypes) {
            $issues += "InterfaceType contains invalid value '$it' (must be Any/Wired/Wireless/RemoteAccess)"
        }
    }

    return $issues
}

# ========================================
# Helper: ConvertTo-RuleParams
# ========================================
# Build the splat hash for New-NetFirewallRule from a CSV row. Empty
# optional cells are omitted so the cmdlet's defaults apply.
function ConvertTo-RuleParams {
    param($Row)

    $params = @{
        DisplayName = $Row.DisplayName.Trim()
        Direction   = $Row.Direction.Trim()
        Action      = $Row.Action.Trim()
    }

    if ($Row.Name)                { $params.Name                = $Row.Name.Trim() }
    if ($Row.RuleEnabled)         { $params.Enabled             = $Row.RuleEnabled.Trim() }
    if ($Row.Description)         { $params.Description         = $Row.Description }
    if ($Row.Group)               { $params.Group               = $Row.Group }
    if ($Row.Program)             { $params.Program             = $Row.Program }
    if ($Row.Service)             { $params.Service             = $Row.Service }
    if ($Row.Protocol)            { $params.Protocol            = $Row.Protocol.Trim() }
    if ($Row.EdgeTraversalPolicy) { $params.EdgeTraversalPolicy = $Row.EdgeTraversalPolicy.Trim() }
    if ($Row.LocalUser)           { $params.LocalUser           = $Row.LocalUser }
    if ($Row.RemoteUser)          { $params.RemoteUser          = $Row.RemoteUser }
    if ($Row.RemoteMachine)       { $params.RemoteMachine       = $Row.RemoteMachine }

    foreach ($pair in @(
        @{ Csv = 'Profile';        Param = 'Profile' }
        @{ Csv = 'LocalPort';      Param = 'LocalPort' }
        @{ Csv = 'RemotePort';     Param = 'RemotePort' }
        @{ Csv = 'LocalAddress';   Param = 'LocalAddress' }
        @{ Csv = 'RemoteAddress';  Param = 'RemoteAddress' }
        @{ Csv = 'IcmpType';       Param = 'IcmpType' }
        @{ Csv = 'InterfaceType';  Param = 'InterfaceType' }
        @{ Csv = 'InterfaceAlias'; Param = 'InterfaceAlias' }
    )) {
        $values = Split-MultiValue -Value $Row.($pair.Csv)
        if ($values.Count -gt 0) {
            $params[$pair.Param] = $values
        }
    }

    return $params
}

# ========================================
# Helper: Test-CreatedRule
# ========================================
# Compare key attributes of a created rule against the source CSV row.
# Returns array of issue strings (empty = match).
function Test-CreatedRule {
    param($Row, $CreatedRule)

    if ($null -eq $CreatedRule) {
        return @("rule not found after creation")
    }

    $issues = @()

    if ($Row.Name -and $CreatedRule.Name -ne $Row.Name.Trim()) {
        $issues += "Name mismatch (expected=$($Row.Name), actual=$($CreatedRule.Name))"
    }
    if ($CreatedRule.Direction.ToString() -ne $Row.Direction.Trim()) {
        $issues += "Direction mismatch (csv=$($Row.Direction), actual=$($CreatedRule.Direction))"
    }
    if ($CreatedRule.Action.ToString() -ne $Row.Action.Trim()) {
        $issues += "Action mismatch (csv=$($Row.Action), actual=$($CreatedRule.Action))"
    }

    $expectedEnabled = if ($Row.RuleEnabled -eq 'False') { 'False' } else { 'True' }
    if ($CreatedRule.Enabled.ToString() -ne $expectedEnabled) {
        $issues += "Enabled mismatch (expected=$expectedEnabled, actual=$($CreatedRule.Enabled))"
    }

    $expectedProfiles = Split-MultiValue -Value $Row.Profile
    if ($expectedProfiles.Count -gt 0) {
        $expectedSorted = ($expectedProfiles | Sort-Object) -join ','
        $actualSorted = (($CreatedRule.Profile.ToString() -split ',') | ForEach-Object { $_.Trim() } | Sort-Object) -join ','
        if ($expectedSorted -ne $actualSorted) {
            $issues += "Profile mismatch (expected=$expectedSorted, actual=$actualSorted)"
        }
    }

    if ($Row.EdgeTraversalPolicy) {
        if ($CreatedRule.EdgeTraversalPolicy.ToString() -ne $Row.EdgeTraversalPolicy.Trim()) {
            $issues += "EdgeTraversalPolicy mismatch (expected=$($Row.EdgeTraversalPolicy), actual=$($CreatedRule.EdgeTraversalPolicy))"
        }
    }

    # Note: InterfaceType / InterfaceAlias / LocalUser / RemoteUser / RemoteMachine
    # require separate Get-NetFirewall*Filter cmdlets to read back. Verification
    # for these fields is intentionally skipped in v1; trust the cmdlet not to
    # silently drop them. See Guide.txt.

    return $issues
}

# ========================================
# Step 0: Privilege Check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required for New-NetFirewallRule."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "firewall_rule_make_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "DisplayName", "Direction", "Action")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load firewall_rule_make_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Validate + idempotency check
# ========================================
$plans = @()
foreach ($row in $enabledItems) {
    $validationIssues = Test-RowValidity -Row $row
    if ($validationIssues.Count -gt 0) {
        $plans += [PSCustomObject]@{
            Action         = 'Invalid'
            Row            = $row
            Reasons        = $validationIssues
            ProgramWarning = $null
        }
        continue
    }

    $programWarning = $null
    if ($row.Program -and -not (Test-Path $row.Program)) {
        $programWarning = "Program path not found (rule will still be created): $($row.Program)"
    }

    # Idempotency: skip if a rule with this Name (when specified) OR
    # DisplayName already exists. Name takes precedence because it is a
    # strict identifier in Windows Firewall (DisplayName can collide).
    $existsBy   = $null
    $matchCount = 0

    if ($row.Name) {
        $existingByName = Get-NetFirewallRule -Name $row.Name -ErrorAction SilentlyContinue
        if ($existingByName) {
            $existsBy   = "Name '$($row.Name)'"
            $matchCount = @($existingByName).Count
        }
    }

    if (-not $existsBy) {
        $existingByDisplayName = Get-NetFirewallRule -DisplayName $row.DisplayName -ErrorAction SilentlyContinue
        if ($existingByDisplayName) {
            $existsBy   = "DisplayName '$($row.DisplayName)'"
            $matchCount = @($existingByDisplayName).Count
        }
    }

    if ($existsBy) {
        $plans += [PSCustomObject]@{
            Action         = 'Skip'
            Row            = $row
            Reasons        = @("$existsBy already exists ($matchCount match(es))")
            ProgramWarning = $programWarning
        }
        continue
    }

    $plans += [PSCustomObject]@{
        Action         = 'Create'
        Row            = $row
        Reasons        = @()
        ProgramWarning = $programWarning
    }
}

$createCount  = @($plans | Where-Object { $_.Action -eq 'Create' }).Count
$skipCount    = @($plans | Where-Object { $_.Action -eq 'Skip' }).Count
$invalidCount = @($plans | Where-Object { $_.Action -eq 'Invalid' }).Count

# ========================================
# Step 3: Preview
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Planned Firewall Rules" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($plan in $plans) {
    $row  = $plan.Row
    $name = if ($row.DisplayName) { $row.DisplayName } else { "<no DisplayName>" }

    switch ($plan.Action) {
        'Create' {
            Write-Host "  [CREATE] $name" -ForegroundColor Yellow
        }
        'Skip' {
            Write-Host "  [SKIP - exists] $name" -ForegroundColor Gray
        }
        'Invalid' {
            Write-Host "  [INVALID] $name" -ForegroundColor Red
            foreach ($r in $plan.Reasons) {
                Write-Host "    - $r" -ForegroundColor Red
            }
        }
    }

    if ($plan.Action -in @('Create', 'Skip')) {
        $direction = if ($row.Direction) { $row.Direction } else { '?' }
        $actionVal = if ($row.Action)    { $row.Action }    else { '?' }
        $protocol  = if ($row.Protocol)  { $row.Protocol }  else { 'Any' }
        $localPort = if ($row.LocalPort) { $row.LocalPort } else { 'Any' }
        $profile   = if ($row.Profile)   { $row.Profile }   else { 'Any' }

        Write-Host "    $direction / $actionVal / $protocol / Local=$localPort / Profile=$profile" -ForegroundColor DarkGray

        if ($plan.ProgramWarning) {
            Show-Warning "    $($plan.ProgramWarning)"
        }
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Quick exit: nothing to do (everything is Skip and no Invalid)
if ($createCount -eq 0 -and $invalidCount -eq 0) {
    Show-Info "All $skipCount rule(s) already exist. Nothing to create."
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" `
        -Message "All rules already exist (skipped: $skipCount)" `
        -Verified $true)
}

# Quick exit: only Invalid (nothing to create)
if ($createCount -eq 0) {
    Show-Error "No valid rules to create (skip=$skipCount, invalid=$invalidCount)"
    Write-Host ""
    return (New-BatchResult -Success 0 -Skip $skipCount -Fail $invalidCount `
        -Title "Firewall Rule Maker Results")
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Create the $createCount rule(s) above?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Apply
# ========================================
$successCount = 0
$failCount    = $invalidCount   # invalid rows count toward failures
$createdRules = @()

foreach ($plan in $plans) {
    if ($plan.Action -ne 'Create') { continue }
    $row = $plan.Row

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Creating: $($row.DisplayName)" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        $params  = ConvertTo-RuleParams -Row $row
        $created = New-NetFirewallRule @params -ErrorAction Stop
        Show-Success "Created: $($row.DisplayName) (Name=$($created.Name))"
        $successCount++
        $createdRules += [PSCustomObject]@{
            Row     = $row
            Created = $created
        }
    }
    catch {
        Show-Error "Failed to create '$($row.DisplayName)': $_"
        $failCount++
    }

    Write-Host ""
}

# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
$verified = $null
if ($createdRules.Count -gt 0) {
    Show-Info "Verifying created rules..."
    Write-Host ""

    $verifyPass = 0
    $verifyFail = 0

    foreach ($entry in $createdRules) {
        $row     = $entry.Row
        # Re-fetch by Name (GUID-like) so we read the actual stored state
        $rule    = Get-NetFirewallRule -Name $entry.Created.Name -ErrorAction SilentlyContinue
        $issues  = Test-CreatedRule -Row $row -CreatedRule $rule

        if ($issues.Count -eq 0) {
            Write-Host "  [VERIFIED] $($row.DisplayName)" -ForegroundColor Green
            $verifyPass++
        }
        else {
            Write-Host "  [VERIFY FAILED] $($row.DisplayName) : $($issues -join '; ')" -ForegroundColor Red
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
    -Title "Firewall Rule Maker Results" -Verified $verified)
