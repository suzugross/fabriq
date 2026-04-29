# Phase 4 functional smoke test (not Pester).
# Tests helper functions and reject-mode paths only. Does NOT trigger
# actual hostname or IP mutation - that requires manual interactive
# verification on a real test machine.
# Run: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <this>

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

Write-Host '--- Set-FabriqIosHostEnvironment ---' -ForegroundColor Cyan
$preserved = @{
    NewPC = $env:SELECTED_NEW_PCNAME
    EthIP = $env:SELECTED_ETH_IP
    EthSubnet = $env:SELECTED_ETH_SUBNET
}
$fakeRow = [PSCustomObject]@{
    AdminID         = '999'
    OldPCName       = 'OLD-FAKE'
    NewPCName       = 'NEW-FAKE-01'
    EthernetIP      = '10.0.0.99'
    EthernetSubnet  = '255.255.255.0'
    EthernetGateway = '10.0.0.1'
    WifiIP          = ''
    WifiSubnet      = ''
    WifiGateway     = ''
    DNS1            = '10.0.0.10'
    DNS2            = ''
    DNS3            = ''
    DNS4            = ''
    Pin             = ''
}
Set-FabriqIosHostEnvironment -Row $fakeRow
Check 'SELECTED_NEW_PCNAME set'  ($env:SELECTED_NEW_PCNAME -eq 'NEW-FAKE-01')
Check 'SELECTED_ETH_IP set'      ($env:SELECTED_ETH_IP    -eq '10.0.0.99')
Check 'SELECTED_ETH_SUBNET set'  ($env:SELECTED_ETH_SUBNET -eq '255.255.255.0')
Check 'SELECTED_DNS1 set'        ($env:SELECTED_DNS1      -eq '10.0.0.10')
# Restore
$env:SELECTED_NEW_PCNAME = $preserved.NewPC
$env:SELECTED_ETH_IP     = $preserved.EthIP
$env:SELECTED_ETH_SUBNET = $preserved.EthSubnet

Write-Host '--- Find-HostlistRowByNewName ---' -ForegroundColor Cyan
$state = New-ShellState
$row = Find-HostlistRowByNewName -NewName 'NEW-PC-01' -State $state
Check 'finds NEW-PC-01 in hostlist' ($null -ne $row -and $row.NewPCName -eq 'NEW-PC-01')
$row = Find-HostlistRowByNewName -NewName 'DEFINITELY-NOT-IN-HOSTLIST' -State $state
Check 'returns null for missing NewName' ($null -eq $row)

Write-Host '--- Get-HostnameCompletionFromHostlist ---' -ForegroundColor Cyan
$names = @(Get-HostnameCompletionFromHostlist -State $state)
Check 'returns at least one name' ($names.Count -ge 1)
Check 'includes NEW-PC-01'        ($names -contains 'NEW-PC-01')

Write-Host '--- Get-IpAddressCompletionFromHostlist ---' -ForegroundColor Cyan
$preservedIp = $env:SELECTED_ETH_IP
$env:SELECTED_ETH_IP = '192.168.42.7'
$cands = @(Get-IpAddressCompletionFromHostlist -State $state)
Check 'always includes from-hostlist'  ($cands -contains 'from-hostlist')
Check 'includes current SELECTED_ETH_IP' ($cands -contains '192.168.42.7')
$env:SELECTED_ETH_IP = $preservedIp

Write-Host '--- Get-InterfaceCompletionFromAdapters ---' -ForegroundColor Cyan
$ifaces = @(Get-InterfaceCompletionFromAdapters)
Check 'returns array (possibly empty) without throw' ($null -ne $ifaces)

Write-Host '--- Mutation reject-paths (no actual hostname/IP change) ---' -ForegroundColor Cyan

# Invoke-HostnameSelection rejects in non-GlobalConfig mode
$state = New-ShellState
$state.Mode = 'PrivilegedExec'
$out = Get-CapturedOutput { Invoke-HostnameSelection -NewName 'NEW-PC-01' -State $state }
Check 'hostname rejected outside GlobalConfig' ($out -match 'global configuration mode')

# Invoke-HostnameSelection with unknown NewName -> refused syslog, no module run
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Invoke-HostnameSelection -NewName 'CERTAINLY-NOT-A-REAL-HOST' -State $state }
Check 'unknown NewName -> refused syslog' ($out -match '%FABRIQ-3-HOSTNAME')
Check 'unknown NewName -> friendly hint'  ($out -match 'has no row with NewPCName')

# Invoke-IpAddressFromHostlist rejects in non-InterfaceConfig mode
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Invoke-IpAddressFromHostlist -State $state }
Check 'ip address rejected outside InterfaceConfig' ($out -match 'interface configuration mode')

# Invoke-IpAddressFromHostlist without host context refuses
$state = New-ShellState
$state.Mode = 'InterfaceConfig'
$state.CurrentInterface = 'TestEthernet'
$preservedNew = $env:SELECTED_NEW_PCNAME
$env:SELECTED_NEW_PCNAME = ''
$out = Get-CapturedOutput { Invoke-IpAddressFromHostlist -State $state }
Check 'ip address without host context refuses' ($out -match 'No host context')
$env:SELECTED_NEW_PCNAME = $preservedNew

# Invoke-IpAddressManual rejects in non-InterfaceConfig mode
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Invoke-IpAddressManual -Ip '10.0.0.1' -Mask '255.255.255.0' -State $state }
Check 'ip address manual rejected outside InterfaceConfig' ($out -match 'interface configuration mode')

Write-Host '--- Dispatcher integration (parser -> mode -> stub-or-real) ---' -ForegroundColor Cyan

# 'host NEW-PC-01' in GlobalConfig -> Invoke-HostnameSelection -> tries to find row.
# We feed an unknown name so it short-circuits at the lookup (no module run).
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$resolved = Resolve-FabriqIosCommand -Tokens @('host','UNKNOWN-FOR-SMOKE') -Mode 'GlobalConfig'
Check 'parser expands host -> hostname' ($resolved.Command -eq 'hostname')
$out = Get-CapturedOutput { Invoke-GlobalConfigCommand -Resolved $resolved -State $state }
Check 'dispatcher reaches refused-syslog path' ($out -match '%FABRIQ-3-HOSTNAME')

# 'ip address from-hostlist' in InterfaceConfig with no host context -> short-circuit
$state = New-ShellState
$state.Mode = 'InterfaceConfig'
$state.CurrentInterface = 'X'
$preservedNew = $env:SELECTED_NEW_PCNAME
$env:SELECTED_NEW_PCNAME = ''
$resolved = Resolve-FabriqIosCommand -Tokens @('ip','address','from-hostlist') -Mode 'InterfaceConfig'
Check 'parser ip address from-hostlist Args.Count=2' ($resolved.Args.Count -eq 2)
Check 'parser ip address from-hostlist Args[1]'      ($resolved.Args[1] -eq 'from-hostlist')
$out = Get-CapturedOutput { Invoke-InterfaceConfigCommand -Resolved $resolved -State $state }
Check 'ip address dispatcher hits No host context' ($out -match 'No host context')
$env:SELECTED_NEW_PCNAME = $preservedNew

Write-Host ''
Write-Host ("=== Summary: {0} passed, {1} failed ===" -f $script:Pass, $script:Fail) `
    -ForegroundColor (& { if ($script:Fail -eq 0) { 'Green' } else { 'Red' } })
exit $script:Fail
