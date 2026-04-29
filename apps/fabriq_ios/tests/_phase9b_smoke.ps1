# Phase 9b functional smoke test - JSON object-form entries (multi-script
# / multi-CSV modules). Verifies Resolve-ModuleEntry normalisation, the
# per-entry Dir / Script / Csv overrides, and dispatcher routing for the
# 9 object-form entries declared in module_categories.json:
#   settings : local_user_create, sysprep_main, sysprep_unattend,
#              sysprep_setupcomplete
#   cleanup  : destroy_history, destroy_ssid, local_user_delete
#   install  : printer_driver_install, printer_register
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

Write-Host '--- Resolve-ModuleEntry: string-form (implicit defaults) ---' -ForegroundColor Cyan
$e = Resolve-ModuleEntry -CategoryId 'settings' -Name 'reg_hklm_config'
Check 'string entry resolved'           ($null -ne $e)
Check 'string Name = reg_hklm_config'   ($e.Name -eq 'reg_hklm_config')
Check 'string Dir = name'               ($e.Dir -eq 'reg_hklm_config')
Check 'string Script = name.ps1'        ($e.Script -eq 'reg_hklm_config.ps1')
Check 'string Csv = $null (auto)'       ($null -eq $e.Csv)
Check 'string Label = name'             ($e.Label -eq 'reg_hklm_config')
Check 'string Category = settings'      ($e.Category -eq 'settings')

Write-Host '--- Resolve-ModuleEntry: object-form (multi-script splits) ---' -ForegroundColor Cyan
# local_user_create: settings, dir=local_user_config, script=local_user_config.ps1
$e = Resolve-ModuleEntry -CategoryId 'settings' -Name 'local_user_create'
Check 'local_user_create resolved'      ($null -ne $e)
Check 'local_user_create Dir override'  ($e.Dir -eq 'local_user_config')
Check 'local_user_create Script override' ($e.Script -eq 'local_user_config.ps1')
Check 'local_user_create Csv override'  ($e.Csv -eq 'local_user_list.csv')
Check 'local_user_create Category'      ($e.Category -eq 'settings')

# local_user_delete: cleanup, same dir, different script
$e = Resolve-ModuleEntry -CategoryId 'cleanup' -Name 'local_user_delete'
Check 'local_user_delete resolved'      ($null -ne $e)
Check 'local_user_delete Dir override'  ($e.Dir -eq 'local_user_config')
Check 'local_user_delete Script override' ($e.Script -eq 'local_user_delete.ps1')
Check 'local_user_delete shares CSV'    ($e.Csv -eq 'local_user_list.csv')
Check 'local_user_delete Category'      ($e.Category -eq 'cleanup')

# Sysprep three-way split: same dir + same script, different CSVs
$e = Resolve-ModuleEntry -CategoryId 'settings' -Name 'sysprep_main'
Check 'sysprep_main Csv = sysprep_list' ($e.Csv -eq 'sysprep_list.csv')
$e = Resolve-ModuleEntry -CategoryId 'settings' -Name 'sysprep_unattend'
Check 'sysprep_unattend Csv = unattend_list' ($e.Csv -eq 'unattend_list.csv')
$e = Resolve-ModuleEntry -CategoryId 'settings' -Name 'sysprep_setupcomplete'
Check 'sysprep_setupcomplete Csv = setupcomplete_list' ($e.Csv -eq 'setupcomplete_list.csv')

# History destroyer two-way split
$e = Resolve-ModuleEntry -CategoryId 'cleanup' -Name 'destroy_history'
Check 'destroy_history Dir = history_destroyer' ($e.Dir -eq 'history_destroyer')
Check 'destroy_history Csv = destroy_list'      ($e.Csv -eq 'destroy_list.csv')
$e = Resolve-ModuleEntry -CategoryId 'cleanup' -Name 'destroy_ssid'
Check 'destroy_ssid Csv = ssid_list'    ($e.Csv -eq 'ssid_list.csv')

# Printer driver split: install vs register (different scripts)
$e = Resolve-ModuleEntry -CategoryId 'install' -Name 'printer_driver_install'
Check 'printer_driver_install Dir override'   ($e.Dir -eq 'printer_driver_config')
Check 'printer_driver_install Script override' ($e.Script -eq 'printer_driver_install.ps1')
$e = Resolve-ModuleEntry -CategoryId 'install' -Name 'printer_register'
Check 'printer_register Script override'      ($e.Script -eq 'printer_config.ps1')
Check 'printer_register Csv = printer_list'   ($e.Csv -eq 'printer_list.csv')

Write-Host '--- Resolve-ModuleEntry: misses ---' -ForegroundColor Cyan
Check 'unknown name -> null'            ($null -eq (Resolve-ModuleEntry -CategoryId 'settings' -Name 'no_such_module'))
Check 'unknown category -> null'        ($null -eq (Resolve-ModuleEntry -CategoryId 'no_cat' -Name 'reg_hklm_config'))
Check 'wrong category -> null'          ($null -eq (Resolve-ModuleEntry -CategoryId 'cleanup' -Name 'reg_hklm_config'))
Check 'excluded name -> null'           ($null -eq (Resolve-ModuleEntry -CategoryId 'settings' -Name 'hostname_config'))
Check 'empty name -> null'              ($null -eq (Resolve-ModuleEntry -CategoryId 'settings' -Name ''))

Write-Host '--- Find-ModuleEntryAcrossCategories ---' -ForegroundColor Cyan
$e = Find-ModuleEntryAcrossCategories -Name 'destroy_history'
Check 'cross-cat finds destroy_history'     ($null -ne $e -and $e.Category -eq 'cleanup')
$e = Find-ModuleEntryAcrossCategories -Name 'printer_register'
Check 'cross-cat finds printer_register'    ($null -ne $e -and $e.Category -eq 'install')
$e = Find-ModuleEntryAcrossCategories -Name 'reg_hklm_config'
Check 'cross-cat finds string entry'        ($null -ne $e -and $e.Category -eq 'settings')
Check 'cross-cat unknown -> null'           ($null -eq (Find-ModuleEntryAcrossCategories -Name 'nope_nope'))

Write-Host '--- Find-ModulePath honours object-form overrides ---' -ForegroundColor Cyan
# local_user_create -> local_user_config\local_user_config.ps1
$p = Find-ModulePath -Name 'local_user_create' -CategoryId 'settings'
Check 'local_user_create path resolved' ($null -ne $p)
Check 'local_user_create script name'   ($p -and (Split-Path $p -Leaf) -eq 'local_user_config.ps1')
Check 'local_user_create dir name'      ($p -and (Split-Path (Split-Path $p -Parent) -Leaf) -eq 'local_user_config')

# local_user_delete -> local_user_config\local_user_delete.ps1
$p = Find-ModulePath -Name 'local_user_delete' -CategoryId 'cleanup'
Check 'local_user_delete path resolved' ($null -ne $p)
Check 'local_user_delete script name'   ($p -and (Split-Path $p -Leaf) -eq 'local_user_delete.ps1')

# printer_driver_install -> printer_driver_config\printer_driver_install.ps1
$p = Find-ModulePath -Name 'printer_driver_install' -CategoryId 'install'
Check 'printer_driver_install path'     ($null -ne $p)
Check 'printer_driver_install leaf'     ($p -and (Split-Path $p -Leaf) -eq 'printer_driver_install.ps1')

# printer_register -> printer_driver_config\printer_config.ps1
$p = Find-ModulePath -Name 'printer_register' -CategoryId 'install'
Check 'printer_register path resolved'  ($null -ne $p)
Check 'printer_register script name'    ($p -and (Split-Path $p -Leaf) -eq 'printer_config.ps1')

# String-form fallback still works
$p = Find-ModulePath -Name 'reg_hklm_config' -CategoryId 'settings'
Check 'reg_hklm_config path resolved'   ($null -ne $p)
Check 'reg_hklm_config script name'     ($p -and (Split-Path $p -Leaf) -eq 'reg_hklm_config.ps1')

# CategoryId-less call uses cross-cat search
$p = Find-ModulePath -Name 'destroy_history'
Check 'destroy_history (no cat) resolved' ($null -ne $p)
Check 'destroy_history (no cat) leaf'     ($p -and (Split-Path $p -Leaf) -eq 'history_destroyer.ps1')

Write-Host '--- Get-ModuleCsvSchema honours csv override ---' -ForegroundColor Cyan
$s = Get-ModuleCsvSchema -Name 'sysprep_main' -CategoryId 'settings'
Check 'sysprep_main schema resolved'    ($null -ne $s)
Check 'sysprep_main CsvFileName'        ($s -and $s.CsvFileName -eq 'sysprep_list.csv')
Check 'sysprep_main has columns'        ($s -and $s.Columns.Count -gt 0)
Check 'sysprep_main ScriptPath set'     ($s -and -not [string]::IsNullOrEmpty($s.ScriptPath))

$s = Get-ModuleCsvSchema -Name 'sysprep_unattend' -CategoryId 'settings'
Check 'sysprep_unattend CsvFileName'    ($s -and $s.CsvFileName -eq 'unattend_list.csv')
$s = Get-ModuleCsvSchema -Name 'sysprep_setupcomplete' -CategoryId 'settings'
Check 'sysprep_setupcomplete CsvFileName' ($s -and $s.CsvFileName -eq 'setupcomplete_list.csv')

$s = Get-ModuleCsvSchema -Name 'destroy_ssid' -CategoryId 'cleanup'
Check 'destroy_ssid CsvFileName'        ($s -and $s.CsvFileName -eq 'ssid_list.csv')

$s = Get-ModuleCsvSchema -Name 'local_user_create' -CategoryId 'settings'
Check 'local_user_create CsvFileName'   ($s -and $s.CsvFileName -eq 'local_user_list.csv')

# String-form auto-detects first *_list.csv
$s = Get-ModuleCsvSchema -Name 'reg_hklm_config' -CategoryId 'settings'
Check 'reg_hklm_config schema resolved' ($null -ne $s)
Check 'reg_hklm_config CsvFileName'     ($s -and $s.CsvFileName -match '_list\.csv$')

Write-Host '--- Tab completion exposes object-form entries ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'

$r = Get-FabriqIosCompletion -Line 'cleanup ' -Position 8 -Mode 'GlobalConfig' -State $state
Check 'cleanup includes destroy_history'    ($r -contains 'destroy_history')
Check 'cleanup includes destroy_ssid'       ($r -contains 'destroy_ssid')
Check 'cleanup includes local_user_delete'  ($r -contains 'local_user_delete')

$r = Get-FabriqIosCompletion -Line 'module ' -Position 7 -Mode 'GlobalConfig' -State $state
Check 'module includes local_user_create'   ($r -contains 'local_user_create')
Check 'module includes sysprep_main'        ($r -contains 'sysprep_main')
Check 'module includes sysprep_unattend'    ($r -contains 'sysprep_unattend')
Check 'module includes sysprep_setupcomplete' ($r -contains 'sysprep_setupcomplete')

$r = Get-FabriqIosCompletion -Line 'install ' -Position 8 -Mode 'GlobalConfig' -State $state
Check 'install includes printer_driver_install' ($r -contains 'printer_driver_install')
Check 'install includes printer_register'       ($r -contains 'printer_register')

# Prefix completion narrows to the right entries
$r = Get-FabriqIosCompletion -Line 'module sysprep_' -Position 15 -Mode 'GlobalConfig' -State $state
Check 'sysprep_ prefix yields 3 entries' ($r.Count -eq 3)
Check 'sysprep_ prefix has main'         ($r -contains 'sysprep_main')
Check 'sysprep_ prefix has unattend'     ($r -contains 'sysprep_unattend')
Check 'sysprep_ prefix has setupcomplete' ($r -contains 'sysprep_setupcomplete')

Write-Host '--- Dispatcher: object-form modules enter ModuleConfig ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'cleanup' -Name 'local_user_delete' -State $state | Out-Null
Check 'local_user_delete -> ModuleConfig'   ($state.Mode -eq 'ModuleConfig')
Check 'local_user_delete CurrentCategoryId' ($state.CurrentCategoryId -eq 'cleanup')
Check 'local_user_delete prompt'            ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-clean)#')

$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'module' -Name 'sysprep_main' -State $state | Out-Null
Check 'sysprep_main -> ModuleConfig'    ($state.Mode -eq 'ModuleConfig')
Check 'sysprep_main prompt config-mod'  ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-mod)#')

$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'install' -Name 'printer_register' -State $state | Out-Null
Check 'printer_register -> ModuleConfig' ($state.Mode -eq 'ModuleConfig')
Check 'printer_register prompt'          ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-install)#')

# Wrong-category rejection still works for object-form names
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Enter-CategoryConfigMode -Verb 'module' -Name 'local_user_delete' -State $state }
Check 'module local_user_delete rejected' ($out -match "is not in the 'settings' category")
Check 'state stays GlobalConfig (obj wrong cat)' ($state.Mode -eq 'GlobalConfig')

$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Enter-CategoryConfigMode -Verb 'cleanup' -Name 'sysprep_main' -State $state }
Check 'cleanup sysprep_main rejected'    ($out -match "is not in the 'cleanup' category")

Write-Host ''
Write-Host ("=== Summary: {0} passed, {1} failed ===" -f $script:Pass, $script:Fail) `
    -ForegroundColor (& { if ($script:Fail -eq 0) { 'Green' } else { 'Red' } })
exit $script:Fail
