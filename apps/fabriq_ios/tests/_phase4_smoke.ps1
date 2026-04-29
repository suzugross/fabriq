# Phase 4 functional smoke test (revised for Phase 8).
# Tests helper functions and reject-mode paths only. Does NOT trigger
# actual hostname or IP mutation - that requires manual interactive
# verification on a real test machine.
# Run: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <this>
#
# Phase 8 deletions (no longer tested):
# - Set-FabriqIosHostEnvironment / Find-HostlistRowByNewName /
#   Get-HostnameCompletionFromHostlist / Get-IpAddressCompletionFromHostlist /
#   Invoke-IpAddressFromHostlist
# - All "from-hostlist" semantics
# Replaced with simpler reject-path coverage matching the new
# `hostname <name>` (direct env set) and `ip address <ip> <mask> ...`
# (positional override) commands.

$ErrorActionPreference = 'Stop'

$script:FabriqIosRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$script:FabriqRoot    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path

. (Join-Path $script:FabriqRoot 'kernel\common.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\shell_state.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\parser.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\prompt.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\syslog.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\help.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\dispatch.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\modes\user_exec.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\modes\privileged_exec.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\modes\global_config.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\modes\interface_config.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\show.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\enable_disable.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\hostname.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\interface.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\ip_address.ps1')

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

Write-Host '--- ConvertFrom-SubnetMaskToPrefix ---' -ForegroundColor Cyan
Check '/24 -> 24'           ((ConvertFrom-SubnetMaskToPrefix '255.255.255.0')   -eq 24)
Check '/16 -> 16'           ((ConvertFrom-SubnetMaskToPrefix '255.255.0.0')     -eq 16)
Check '/30 -> 30'           ((ConvertFrom-SubnetMaskToPrefix '255.255.255.252') -eq 30)
Check 'empty -> empty'      ((ConvertFrom-SubnetMaskToPrefix '')                -eq '')
Check 'malformed -> empty'  ((ConvertFrom-SubnetMaskToPrefix '255.255.255')     -eq '')

Write-Host '--- Get-InterfaceCompletionFromAdapters ---' -ForegroundColor Cyan
$ifaces = @(Get-InterfaceCompletionFromAdapters)
Check 'returns array (possibly empty) without throw' ($null -ne $ifaces)

Write-Host '--- Mutation reject-paths (no actual hostname/IP change) ---' -ForegroundColor Cyan

# Invoke-HostnameSelection rejects in non-GlobalConfig mode
$state = New-ShellState
$state.Mode = 'PrivilegedExec'
$out = Get-CapturedOutput { Invoke-HostnameSelection -NewName 'NEW-PC-XX' -State $state }
Check 'hostname rejected outside GlobalConfig' ($out -match 'global configuration mode')

# Invoke-HostnameSelection with empty name
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Invoke-HostnameSelection -NewName '' -State $state }
Check 'empty hostname -> incomplete' ($out -match 'Incomplete')

# Invoke-IpAddressManual rejects in non-InterfaceConfig mode
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Invoke-IpAddressManual -Ip '10.0.0.1' -Mask '255.255.255.0' -State $state }
Check 'ip address manual rejected outside InterfaceConfig' ($out -match 'interface configuration mode')

# Invoke-IpAddressManual without IP/Mask
$state = New-ShellState
$state.Mode = 'InterfaceConfig'
$state.CurrentInterface = 'TestEth'
$out = Get-CapturedOutput { Invoke-IpAddressManual -Ip '' -Mask '' -State $state }
Check 'ip address manual without ip/mask -> Usage' ($out -match 'Usage:')

Write-Host '--- Dispatcher integration (parser -> mode dispatch) ---' -ForegroundColor Cyan

# 'host NEW-PC-XYZ' -> parser expands to 'hostname'
$resolved = Resolve-FabriqIosCommand -Tokens @('host','NEW-PC-XYZ') -Mode 'GlobalConfig'
Check 'parser host -> hostname'      ($resolved.Command -eq 'hostname')
Check 'parser preserves ad-hoc name' ($resolved.Args[0] -eq 'NEW-PC-XYZ')

# `ip address 10.0.0.1 255.255.255.0 10.0.0.1` (4 args after expansion) -> ip
$resolved = Resolve-FabriqIosCommand -Tokens @('ip','address','10.0.0.1','255.255.255.0','10.0.0.1') -Mode 'InterfaceConfig'
Check 'parser ip address resolves'   ($resolved.Command -eq 'ip')
Check 'parser ip address Args.Count=4' ($resolved.Args.Count -eq 4)
Check 'parser ip address arg[0]=address' ($resolved.Args[0] -eq 'address')
Check 'parser ip address arg[3]=gateway' ($resolved.Args[3] -eq '10.0.0.1')

# 5-arg fully-self-contained form (with DNS)
$resolved = Resolve-FabriqIosCommand -Tokens @('ip','address','10.0.0.1','255.255.255.0','10.0.0.1','8.8.8.8','1.1.1.1') -Mode 'InterfaceConfig'
Check 'parser ip address 6 args (full form)' ($resolved.Args.Count -eq 6)

Write-Host ''
Write-Host ("=== Summary: {0} passed, {1} failed ===" -f $script:Pass, $script:Fail) `
    -ForegroundColor (& { if ($script:Fail -eq 0) { 'Green' } else { 'Red' } })
exit $script:Fail
