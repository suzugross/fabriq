# Phase 6 functional smoke test - `module <name>` command in
# GlobalConfig. Phase 7 redesigned the semantics: this command now
# enters ModuleConfig mode (config-mod)# rather than running the
# module immediately. Phase 7 smoke covers the (config-mod)# mode
# itself; this file keeps the GlobalConfig-level invariants
# (auto-discovery, exclusion list, schema lookup, mode entry).
# Run: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <this>

$ErrorActionPreference = 'Stop'

$script:FabriqIosRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$script:FabriqRoot    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path

. (Join-Path $script:FabriqRoot 'kernel\common.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\shell_state.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\parser.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\syslog.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\help.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\dispatch.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\hostname.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\interface.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\ip_address.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\module.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\modes\global_config.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\modes\module_config.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\completer.ps1')

$script:Pass = 0
$script:Fail = 0
function Check($name, $cond) {
    if ($cond) {
        Write-Host ("  PASS  {0}" -f $name) -ForegroundColor Green
        $script:Pass++
    } else {
        Write-Host ("  FAIL  {0}" -f $name) -ForegroundColor Red
        $script:Fail++
    }
}
function Get-CapturedOutput {
    param([scriptblock]$Block)
    $all = & $Block *>&1
    return ($all | Out-String)
}

Write-Host '--- GlobalConfig vocabulary ---' -ForegroundColor Cyan
$vocab = Get-FabriqIosCommandVocabulary -Mode 'GlobalConfig'
Check 'vocab contains module'    ($vocab -contains 'module')
Check 'vocab contains hostname'  ($vocab -contains 'hostname')
Check 'vocab contains interface' ($vocab -contains 'interface')

Write-Host '--- Get-ModuleCompletionFromFilesystem ---' -ForegroundColor Cyan
$names = Get-ModuleCompletionFromFilesystem
Check 'returns at least 50 names' ($names.Count -ge 50)
Check 'sorted alphabetically'     ($names[0] -lt $names[-1])

Check 'excludes windows_update'      (-not ($names -contains 'windows_update'))
Check 'excludes test_error_module'   (-not ($names -contains 'test_error_module'))
Check 'excludes test_harness_config' (-not ($names -contains 'test_harness_config'))
Check 'excludes fabriq_app_launcher' (-not ($names -contains 'fabriq_app_launcher'))
Check 'excludes hostname_config'     (-not ($names -contains 'hostname_config'))
Check 'excludes ipaddress_config'    (-not ($names -contains 'ipaddress_config'))

Check 'includes reg_hklm_config'    ($names -contains 'reg_hklm_config')
Check 'includes bitlocker_config'   ($names -contains 'bitlocker_config')
Check 'includes bloatware_remove'   ($names -contains 'bloatware_remove')
Check 'includes evidence_config'    ($names -contains 'evidence_config')

Write-Host '--- Find-ModulePath ---' -ForegroundColor Cyan
$p = Find-ModulePath -Name 'reg_hklm_config'
Check 'reg_hklm_config resolves'    ($null -ne $p -and (Test-Path $p))
$p = Find-ModulePath -Name 'log_uploader'
Check 'log_uploader (extended) resolves' ($null -ne $p -and (Test-Path $p))
Check 'unknown returns null'        ($null -eq (Find-ModulePath -Name 'definitely-not-a-module'))
Check 'excluded returns null'       ($null -eq (Find-ModulePath -Name 'hostname_config'))
Check 'empty returns null'          ($null -eq (Find-ModulePath -Name ''))

Write-Host '--- Tab/?: completion via Get-FabriqIosCompletion ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'

$r = Get-FabriqIosCompletion -Line 'module ' -Position 7 -Mode 'GlobalConfig' -State $state
Check 'module <space> -> module list (>=50)' ($r.Count -ge 50)
Check 'module <space> -> contains reg_hklm_config' ($r -contains 'reg_hklm_config')
Check 'module <space> -> excludes hostname_config' (-not ($r -contains 'hostname_config'))

$r = Get-FabriqIosCompletion -Line 'module reg_h' -Position 12 -Mode 'GlobalConfig' -State $state
Check 'module reg_h prefix filter -> reg_hklm_config + reg_hkcu_config' ($r.Count -eq 2)
Check 'module reg_h -> contains reg_hklm_config'    ($r -contains 'reg_hklm_config')
Check 'module reg_h -> contains reg_hkcu_config'    ($r -contains 'reg_hkcu_config')

$expanded = Expand-FabriqIosAbbreviation -Tokens @('mod','reg_hklm_config') -Mode 'GlobalConfig'
Check 'mod -> module abbreviation'  ($expanded[0] -eq 'module')
Check 'mod arg preserved'           ($expanded[1] -eq 'reg_hklm_config')

Write-Host '--- Module-entry reject paths (no actual run) ---' -ForegroundColor Cyan

# Enter-ModuleConfigMode outside GlobalConfig is rejected
$state = New-ShellState
$state.Mode = 'PrivilegedExec'
$out = Get-CapturedOutput { Enter-ModuleConfigMode -Name 'reg_hklm_config' -State $state }
Check 'module entry rejected outside GlobalConfig' ($out -match 'global configuration mode')
Check 'state stays PrivilegedExec'                 ($state.Mode -eq 'PrivilegedExec')

# Empty / whitespace name
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Enter-ModuleConfigMode -Name '' -State $state }
Check 'empty name -> incomplete'    ($out -match 'Incomplete')
Check 'state stays GlobalConfig'    ($state.Mode -eq 'GlobalConfig')

# Unknown name
$out = Get-CapturedOutput { Enter-ModuleConfigMode -Name 'definitely-not-a-module' -State $state }
Check 'unknown name -> not found'   ($out -match 'not found')
Check 'state stays GlobalConfig (unknown)' ($state.Mode -eq 'GlobalConfig')

# Excluded name
$out = Get-CapturedOutput { Enter-ModuleConfigMode -Name 'hostname_config' -State $state }
Check 'excluded name -> not found'  ($out -match 'not found')

Write-Host '--- Dispatcher integration (parser -> mode -> reject) ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'

# 'module' alone -> incomplete
$resolved = Resolve-FabriqIosCommand -Tokens @('module') -Mode 'GlobalConfig'
Check 'parser: module alone resolves' ($resolved.Command -eq 'module')
$out = Get-CapturedOutput { Invoke-GlobalConfigCommand -Resolved $resolved -State $state }
Check 'dispatcher: module alone -> incomplete' ($out -match 'Incomplete')

# 'module unknown-x' -> not found
$resolved = Resolve-FabriqIosCommand -Tokens @('module','unknown-for-smoke') -Mode 'GlobalConfig'
Check 'parser: module + arg resolves' ($resolved.Args.Count -eq 1)
$out = Get-CapturedOutput { Invoke-GlobalConfigCommand -Resolved $resolved -State $state }
Check 'dispatcher: unknown -> not found' ($out -match 'not found')
Check 'state stays GlobalConfig (dispatcher)' ($state.Mode -eq 'GlobalConfig')

Write-Host ''
Write-Host ("=== Summary: {0} passed, {1} failed ===" -f $script:Pass, $script:Fail) `
    -ForegroundColor (& { if ($script:Fail -eq 0) { 'Green' } else { 'Red' } })
exit $script:Fail
