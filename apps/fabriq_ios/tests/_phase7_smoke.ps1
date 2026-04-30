# Phase 7 functional smoke test - ModuleConfig mode (config-mod)#
# with ephemeral row injection via Import-ModuleCsv override.
# The actual module-execution leg uses real fabriq modules and
# real OS state, so live runs are validated manually. This smoke
# covers everything around it: schema lookup, mode transitions,
# parser integration, pair parsing, reject paths, and the override
# mechanism with a synthetic test scriptblock.
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
. (Join-Path $script:FabriqIosRoot 'lib\commands\hostname.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\interface.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\ip_address.ps1')
. (Join-Path $script:FabriqIosRoot 'lib\commands\categories.ps1')
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

Write-Host '--- ShellState extensions ---' -ForegroundColor Cyan
$state = New-ShellState
Check 'ConfigModuleName field exists'   ($state.ContainsKey('ConfigModuleName'))
Check 'ConfigModuleSchema field exists' ($state.ContainsKey('ConfigModuleSchema'))

Write-Host '--- Mode transitions ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Set-ShellMode -State $state -NewMode 'ModuleConfig'
Check 'GlobalConfig -> ModuleConfig allowed' ($state.Mode -eq 'ModuleConfig')

# ModuleConfig -> GlobalConfig allowed (exit)
Set-ShellMode -State $state -NewMode 'GlobalConfig'
Check 'ModuleConfig -> GlobalConfig allowed (exit)' ($state.Mode -eq 'GlobalConfig')

# ModuleConfig -> PrivilegedExec allowed (end)
$state.Mode = 'ModuleConfig'
Set-ShellMode -State $state -NewMode 'PrivilegedExec'
Check 'ModuleConfig -> PrivilegedExec allowed (end)' ($state.Mode -eq 'PrivilegedExec')

# UserExec -> ModuleConfig disallowed
$state.Mode = 'UserExec'
$threw = $false
try { Set-ShellMode -State $state -NewMode 'ModuleConfig' } catch { $threw = $true }
Check 'UserExec -> ModuleConfig disallowed' $threw

Write-Host '--- ModuleConfig vocabulary + prompt ---' -ForegroundColor Cyan
$vocab = Get-FabriqIosCommandVocabulary -Mode 'ModuleConfig'
Check 'vocab contains set'  ($vocab -contains 'set')
Check 'vocab contains add'  ($vocab -contains 'add')
Check 'vocab contains show' ($vocab -contains 'show')
Check 'vocab contains exit' ($vocab -contains 'exit')
Check 'vocab contains end'  ($vocab -contains 'end')

$state = New-ShellState
$state.Mode = 'ModuleConfig'
Check 'prompt is fabriq(config-mod)#' ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-mod)#')

Write-Host '--- Get-ModuleCsvSchema ---' -ForegroundColor Cyan
# reg_hklm_config has _list.csv with known columns
$schema = Get-ModuleCsvSchema -Name 'reg_hklm_config'
Check 'reg_hklm_config schema loaded'    ($null -ne $schema)
Check 'reg_hklm_config schema has Columns' ($schema.Columns.Count -ge 3)

# Modules without _list.csv (e.g. excluded ones) return null
$schema = Get-ModuleCsvSchema -Name 'hostname_config'   # excluded
Check 'excluded module schema is null' ($null -eq $schema)

$schema = Get-ModuleCsvSchema -Name 'definitely-not-a-module'
Check 'unknown module schema is null' ($null -eq $schema)

Write-Host '--- Enter-ModuleConfigMode (real entry path) ---' -ForegroundColor Cyan
# Successful entry: reg_hklm_config has both .ps1 and _list.csv
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Enter-CategoryConfigMode -Verb 'module' -Name 'reg_hklm_config' -State $state }
Check 'enter reg_hklm_config -> ModuleConfig'  ($state.Mode -eq 'ModuleConfig')
Check 'enter reg_hklm_config -> ConfigModuleName set' ($state.ConfigModuleName -eq 'reg_hklm_config')
Check 'enter reg_hklm_config -> ConfigModuleSchema set' ($null -ne $state.ConfigModuleSchema)
Check 'enter prints intro' ($out -match 'Configuring module')

Write-Host '--- exit / end transitions clear context ---' -ForegroundColor Cyan
$resolved = @{ Command = 'exit'; Args = @(); Error = $null }
Invoke-ModuleConfigCommand -Resolved $resolved -State $state | Out-Null
Check 'exit -> GlobalConfig'                ($state.Mode -eq 'GlobalConfig')
Check 'exit clears ConfigModuleName'        ($null -eq $state.ConfigModuleName)
Check 'exit clears ConfigModuleSchema'      ($null -eq $state.ConfigModuleSchema)

# end from ModuleConfig
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'module' -Name 'reg_hklm_config' -State $state | Out-Null
$resolved = @{ Command = 'end'; Args = @(); Error = $null }
Invoke-ModuleConfigCommand -Resolved $resolved -State $state | Out-Null
Check 'end -> PrivilegedExec'                ($state.Mode -eq 'PrivilegedExec')
Check 'end clears ConfigModuleName'          ($null -eq $state.ConfigModuleName)

Write-Host '--- show command renders schema ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'module' -Name 'reg_hklm_config' -State $state | Out-Null
$out = Get-CapturedOutput { Show-ModuleConfigSchema -State $state }
Check 'show emits Module: header'   ($out -match 'Module:')
Check 'show emits Columns header'   ($out -match 'Columns:')
Check 'show emits Usage header'     ($out -match 'Usage:')

Write-Host '--- Pair-parsing reject paths (no actual run) ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'module' -Name 'reg_hklm_config' -State $state | Out-Null

# Empty pairs
$out = Get-CapturedOutput { Invoke-ModuleEphemeralRun -PairArgs @() -State $state }
Check 'empty args -> Usage hint' ($out -match 'Usage:')

# Odd pair count
$out = Get-CapturedOutput { Invoke-ModuleEphemeralRun -PairArgs @('Path','HKLM:\Foo','Name') -State $state }
Check 'odd args -> error' ($out -match 'Odd argument count')

# Unknown column
$out = Get-CapturedOutput { Invoke-ModuleEphemeralRun -PairArgs @('NotAColumn','x') -State $state }
Check 'unknown column -> error' ($out -match "Unknown column 'NotAColumn'")

Write-Host '--- Import-ModuleCsv override mechanics (direct) ---' -ForegroundColor Cyan
# Tests the path-matched override + closure + restore mechanics
# directly at the same scope, decoupled from Invoke-FabriqIosModule
# dispatch. Production dispatch follows an `& $modulePath` child-
# scope chain that resolves the overridden Import-ModuleCsv
# correctly; that path is validated by manual run on a real module.
# Here we exercise the override and pass-through routing in
# isolation so the test does not depend on PowerShell's dynamic-
# scope rules for indirect function calls.

$ephemeralRow = [PSCustomObject][ordered]@{
    Path     = 'HKLM:\PROBE'
    Name     = 'Synthetic'
    Type     = 'REG_DWORD'
    Value    = '42'
    Enabled  = '1'
}
$expectedFile = 'reg_hklm_list.csv'
$original = Get-Item Function:Import-ModuleCsv
$originalScript = $original.ScriptBlock

$override = {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$FilterEnabled,
        [string[]]$RequiredColumns,
        [string]$Segment = $env:FABRIQ_SEGMENT
    )
    if ([System.IO.Path]::GetFileName($Path) -eq $expectedFile) {
        if ($FilterEnabled) {
            return @($ephemeralRow | Where-Object { $_.Enabled -eq '1' })
        }
        return @($ephemeralRow)
    }
    return & $originalScript @PSBoundParameters
}.GetNewClosure()

Set-Item Function:Import-ModuleCsv -Value $override
try {
    # Probe the target CSV: any path whose filename matches
    # reg_hklm_list.csv must yield the ephemeral row. Use @() to
    # wrap because PowerShell auto-unwraps single-element returns
    # (matching the existing Import-ModuleCsv contract: caller
    # may receive a scalar or array depending on row count). PS 5.1
    # does not promote PSCustomObject.Count to 1, so wrap explicitly.
    $rows1 = @(Import-ModuleCsv -Path 'C:\arbitrary\reg_hklm_list.csv')
    Check 'override matched filename'        ($rows1.Count -eq 1)
    Check 'ephemeral row Path field'         ($rows1[0].Path -eq 'HKLM:\PROBE')
    Check 'ephemeral row Name field'         ($rows1[0].Name -eq 'Synthetic')
    Check 'ephemeral row Value field'        ($rows1[0].Value -eq '42')
    Check 'ephemeral row Enabled field'      ($rows1[0].Enabled -eq '1')

    # FilterEnabled honoured: row has Enabled=1 -> still returned.
    $rows2 = @(Import-ModuleCsv -Path 'C:\arbitrary\reg_hklm_list.csv' -FilterEnabled)
    Check 'FilterEnabled honoured (Enabled=1)' ($rows2.Count -eq 1)

    # Probe an unrelated CSV: must pass through to the original.
    $hostlist = Join-Path $script:FabriqRoot 'kernel\csv\hostlist.csv'
    $rows3 = @(Import-ModuleCsv -Path $hostlist)
    Check 'unrelated CSV passes through' ($rows3.Count -gt 0)
} finally {
    Set-Item Function:Import-ModuleCsv -Value $originalScript
}

Check 'Import-ModuleCsv restored after run' (
    (Get-Item Function:Import-ModuleCsv).ScriptBlock.ToString() -eq $originalScript.ToString()
)

Write-Host '--- Parser integration ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'

# Full path: parse 'module reg_hklm_config' then dispatch
$tokens = ConvertTo-FabriqIosTokens 'module reg_hklm_config'
$resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode 'GlobalConfig'
Check 'parser: module + reg_hklm_config' ($resolved.Command -eq 'module' -and $resolved.Args[0] -eq 'reg_hklm_config')
Invoke-GlobalConfigCommand -Resolved $resolved -State $state | Out-Null
Check 'dispatcher transitions to ModuleConfig' ($state.Mode -eq 'ModuleConfig')

# Inside ModuleConfig: 'sh' should expand to 'show', not 'set' (exact prefix)
$tokens = ConvertTo-FabriqIosTokens 'sh'
$resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode 'ModuleConfig'
Check 'sh -> show in ModuleConfig' ($resolved.Command -eq 'show')

# 'se' -> 'set'
$tokens = ConvertTo-FabriqIosTokens 'se Path HKLM:\X'
$resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode 'ModuleConfig'
Check 'se -> set abbreviation'  ($resolved.Command -eq 'set')
Check 'se preserves args'       ($resolved.Args.Count -eq 2)

# 'a' -> 'add' (unique in ModuleConfig)
$tokens = ConvertTo-FabriqIosTokens 'a Name foo'
$resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode 'ModuleConfig'
Check 'a -> add abbreviation'   ($resolved.Command -eq 'add')

# 'ex' -> exit, 'en' -> end (not ambiguous in ModuleConfig)
$tokens = ConvertTo-FabriqIosTokens 'ex'
$resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode 'ModuleConfig'
Check 'ex -> exit' ($resolved.Command -eq 'exit')
$tokens = ConvertTo-FabriqIosTokens 'en'
$resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode 'ModuleConfig'
Check 'en -> end'  ($resolved.Command -eq 'end')

Write-Host '--- Tab completion: set/add return column names ---' -ForegroundColor Cyan
# Simulate the global state container used by the PSReadLine handlers
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'module' -Name 'reg_hklm_config' -State $state | Out-Null
$global:_FabriqIosShellState = $state

$candidates = Get-FabriqIosCompletion -Line 'set ' -Position 4 -Mode 'ModuleConfig' -State $state
Check 'set <space> -> column names' ($candidates.Count -ge 3)
$candidates = Get-FabriqIosCompletion -Line 'add ' -Position 4 -Mode 'ModuleConfig' -State $state
Check 'add <space> -> column names' ($candidates.Count -ge 3)

Write-Host '--- Tab completion: set/add multi-pair (Phase 9c) ---' -ForegroundColor Cyan
# After a column with declared enums, value position offers enum values.
$line = 'set Enabled '
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'set Enabled <space> -> enum values'   ($candidates.Count -ge 2)
Check 'set Enabled <space> contains 1'       ($candidates -contains '1')
Check 'set Enabled <space> contains 0'       ($candidates -contains '0')

# Prefix narrowing on enum values.
$line = 'set Type REG_D'
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'set Type REG_D -> filtered enum'      ($candidates -contains 'REG_DWORD')
Check 'set Type REG_D excludes REG_SZ'       (-not ($candidates -contains 'REG_SZ'))

# After col-value pair, next position is a column name again.
$line = 'set Enabled 1 '
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'set Enabled 1 <space> -> column names' ($candidates.Count -ge 3)
Check 'set Enabled 1 <space> contains AdminID' ($candidates -contains 'AdminID')
Check 'set Enabled 1 <space> excludes used Enabled' (-not ($candidates -contains 'Enabled'))

# Prefix narrowing on second column name.
$line = 'set Enabled 1 Ad'
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'set Enabled 1 Ad -> AdminID'          ($candidates -contains 'AdminID')
Check 'set Enabled 1 Ad excludes Enabled'    (-not ($candidates -contains 'Enabled'))

# Three-pair chain: after two pairs, columns come back, used cols hidden.
$line = 'set Enabled 1 Type REG_DWORD '
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check '3rd-pair column completion fires'     ($candidates.Count -ge 1)
Check '3rd-pair excludes Enabled'            (-not ($candidates -contains 'Enabled'))
Check '3rd-pair excludes Type'               (-not ($candidates -contains 'Type'))
Check '3rd-pair includes Value'              ($candidates -contains 'Value')

# add verb gets the same treatment.
$line = 'add Enabled '
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'add Enabled <space> -> enum values'   ($candidates.Count -ge 2)
$line = 'add Enabled 1 '
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'add Enabled 1 <space> -> columns'     ($candidates.Count -ge 3)

# Column with no preset enums returns no value candidates (silent, not error).
$line = 'set SettingTitle '
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'set <enum-less col> <space> -> empty' ($candidates.Count -eq 0)

# Case-insensitive used-column filtering.
$line = 'set enabled 1 '
$candidates = Get-FabriqIosCompletion -Line $line -Position $line.Length -Mode 'ModuleConfig' -State $state
Check 'used-column filter is case-insensitive' (-not ($candidates -contains 'Enabled'))

$global:_FabriqIosShellState = $null

Write-Host ''
Write-Host ("=== Summary: {0} passed, {1} failed ===" -f $script:Pass, $script:Fail) `
    -ForegroundColor (& { if ($script:Fail -eq 0) { 'Green' } else { 'Red' } })
exit $script:Fail
