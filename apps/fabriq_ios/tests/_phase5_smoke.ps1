# Phase 5 functional smoke test.
# Tests the pure completion engine. PSReadLine integration itself
# requires interactive verification (manual test).
# Run: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <this>

$ErrorActionPreference = 'Stop'

$script:FabriqIosRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$script:FabriqRoot    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path

. (Join-Path $script:FabriqRoot 'kernel\common.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\shell_state.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\parser.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\dispatch.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\hostname.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\interface.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\ip_address.ps1')
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

$state = New-ShellState

Write-Host '--- Get-FabriqIosCompletion: top-level vocabulary ---' -ForegroundColor Cyan
$r = Get-FabriqIosCompletion -Line '' -Position 0 -Mode 'UserExec' -State $state
Check 'UserExec empty -> contains enable'  ($r -contains 'enable')
Check 'UserExec empty -> contains show'    ($r -contains 'show')
Check 'UserExec empty -> contains exit'    ($r -contains 'exit')

$r = Get-FabriqIosCompletion -Line '' -Position 0 -Mode 'PrivilegedExec' -State $state
Check 'PrivilegedExec empty -> contains configure' ($r -contains 'configure')
Check 'PrivilegedExec empty -> contains reload'    ($r -contains 'reload')

$r = Get-FabriqIosCompletion -Line '' -Position 0 -Mode 'GlobalConfig' -State $state
Check 'GlobalConfig empty -> contains hostname'  ($r -contains 'hostname')
Check 'GlobalConfig empty -> contains interface' ($r -contains 'interface')

$r = Get-FabriqIosCompletion -Line '' -Position 0 -Mode 'InterfaceConfig' -State $state
Check 'InterfaceConfig empty -> contains ip'  ($r -contains 'ip')
Check 'InterfaceConfig empty -> contains end' ($r -contains 'end')

Write-Host '--- Get-FabriqIosCompletion: prefix filtering ---' -ForegroundColor Cyan
$r = Get-FabriqIosCompletion -Line 'e' -Position 1 -Mode 'UserExec' -State $state
Check 'UserExec e -> enable'  ($r -contains 'enable')
Check 'UserExec e -> exit'    ($r -contains 'exit')
Check 'UserExec e -> NOT show' (-not ($r -contains 'show'))

$r = Get-FabriqIosCompletion -Line 'sho' -Position 3 -Mode 'PrivilegedExec' -State $state
Check 'sho -> show'      ($r -contains 'show')
Check 'sho count = 1'    ($r.Count -eq 1)

Write-Host '--- Get-FabriqIosCompletion: subcommand expansion ---' -ForegroundColor Cyan
$r = Get-FabriqIosCompletion -Line 'show ' -Position 5 -Mode 'UserExec' -State $state
Check 'show <space> UserExec -> version'   ($r -contains 'version')
Check 'show <space> UserExec -> manifesto' ($r -contains 'manifesto')
Check 'show <space> UserExec -> NOT running-config' (-not ($r -contains 'running-config'))

$r = Get-FabriqIosCompletion -Line 'show ' -Position 5 -Mode 'PrivilegedExec' -State $state
Check 'show <space> Priv -> running-config' ($r -contains 'running-config')
Check 'show <space> Priv -> profiles'       ($r -contains 'profiles')
Check 'show <space> Priv -> evidence'       ($r -contains 'evidence')

$r = Get-FabriqIosCompletion -Line 'show ru' -Position 7 -Mode 'PrivilegedExec' -State $state
Check 'show ru -> running-config' ($r -contains 'running-config')
Check 'show ru count = 1'         ($r.Count -eq 1)

$r = Get-FabriqIosCompletion -Line 'show ho' -Position 7 -Mode 'PrivilegedExec' -State $state
Check 'show ho -> host'  ($r -contains 'host')
Check 'show ho -> hosts' ($r -contains 'hosts')
Check 'show ho count >= 2' ($r.Count -ge 2)

Write-Host '--- Get-FabriqIosCompletion: dynamic sources ---' -ForegroundColor Cyan
$r = Get-FabriqIosCompletion -Line 'host ' -Position 5 -Mode 'GlobalConfig' -State $state
Check 'host <space> GlobalConfig -> hostlist NewPC' ($r -contains 'NEW-PC-01')

$r = Get-FabriqIosCompletion -Line 'host NEW-PC-' -Position 12 -Mode 'GlobalConfig' -State $state
$allStartWithPrefix = $true
foreach ($item in $r) { if (-not $item.StartsWith('NEW-PC-')) { $allStartWithPrefix = $false } }
Check 'host NEW-PC- prefix filter' ($r.Count -ge 1 -and $allStartWithPrefix)

$r = Get-FabriqIosCompletion -Line 'ip address ' -Position 11 -Mode 'InterfaceConfig' -State $state
Check 'ip address <space> -> from-hostlist' ($r -contains 'from-hostlist')

$r = Get-FabriqIosCompletion -Line 'ip add' -Position 6 -Mode 'InterfaceConfig' -State $state
Check 'ip add -> address' ($r -contains 'address')

$r = Get-FabriqIosCompletion -Line 'interface ' -Position 10 -Mode 'GlobalConfig' -State $state
Check 'interface <space> returns array (no throw)' ($null -ne $r)

Write-Host '--- Edge cases ---' -ForegroundColor Cyan
$r = Get-FabriqIosCompletion -Line 'nonsense ' -Position 9 -Mode 'UserExec' -State $state
Check 'unknown command -> empty'  ($r.Count -eq 0)

$r = Get-FabriqIosCompletion -Line '   ' -Position 3 -Mode 'UserExec' -State $state
Check 'whitespace-only -> top vocab' ($r -contains 'enable')

$r = Get-FabriqIosCompletion -Line 'x' -Position 1 -Mode 'UserExec' -State $state
Check 'no-prefix-match -> empty' ($r.Count -eq 0)

Write-Host '--- Get-CommonPrefix ---' -ForegroundColor Cyan
Check 'hello/help/helpful -> hel' ((Get-CommonPrefix -Strings @('hello','help','helpful')) -eq 'hel')
Check 'foo/bar -> empty'          ((Get-CommonPrefix -Strings @('foo','bar')) -eq '')
Check 'single string passthrough' ((Get-CommonPrefix -Strings @('alone')) -eq 'alone')
Check 'empty input -> empty'      ((Get-CommonPrefix -Strings @()) -eq '')
Check 'case-insensitive Hel'      ((Get-CommonPrefix -Strings @('Hello','HELP')) -eq 'Hel')

Write-Host ''
Write-Host ("=== Summary: {0} passed, {1} failed ===" -f $script:Pass, $script:Fail) `
    -ForegroundColor (& { if ($script:Fail -eq 0) { 'Green' } else { 'Red' } })
exit $script:Fail
