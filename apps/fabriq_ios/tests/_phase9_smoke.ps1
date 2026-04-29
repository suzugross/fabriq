# Phase 9 functional smoke test - category-driven verb dispatch.
# Tests JSON loading, category lookup, verb-to-category routing,
# mode entry per category (with prompt-suffix verification), wrong-
# category rejection, and the new completer routes. Real module
# execution is still manual.
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

Write-Host '--- Get-FabriqIosCategories (JSON load) ---' -ForegroundColor Cyan
$cats = Get-FabriqIosCategories
Check 'returns hashtable'             ($cats -is [hashtable])
Check 'has settings category'         ($cats.ContainsKey('settings'))
Check 'has cleanup category'          ($cats.ContainsKey('cleanup'))
Check 'has copy category'             ($cats.ContainsKey('copy'))
Check 'has install category'          ($cats.ContainsKey('install'))
Check 'has scripting category'        ($cats.ContainsKey('scripting'))
Check 'all 5 categories'              ($cats.Count -eq 5)

Write-Host '--- Get-FabriqIosCategoryByVerb ---' -ForegroundColor Cyan
$cat = Get-FabriqIosCategoryByVerb -Verb 'module'
Check 'verb module -> settings'        ($cat.id -eq 'settings')
$cat = Get-FabriqIosCategoryByVerb -Verb 'cleanup'
Check 'verb cleanup -> cleanup'        ($cat.id -eq 'cleanup')
$cat = Get-FabriqIosCategoryByVerb -Verb 'copy'
Check 'verb copy -> copy'              ($cat.id -eq 'copy')
$cat = Get-FabriqIosCategoryByVerb -Verb 'install'
Check 'verb install -> install'        ($cat.id -eq 'install')
$cat = Get-FabriqIosCategoryByVerb -Verb 'script'
Check 'verb script -> scripting'       ($cat.id -eq 'scripting')
Check 'unknown verb -> null'           ($null -eq (Get-FabriqIosCategoryByVerb -Verb 'nope'))

Write-Host '--- Get-CategoryModuleCompletion (counts + exclusions) ---' -ForegroundColor Cyan
$settings = Get-CategoryModuleCompletion -CategoryId 'settings'
Check 'settings has many modules'      ($settings.Count -gt 30)
Check 'settings includes reg_hklm_config' ($settings -contains 'reg_hklm_config')
Check 'settings excludes hostname_config' (-not ($settings -contains 'hostname_config'))

$cleanup = Get-CategoryModuleCompletion -CategoryId 'cleanup'
Check 'cleanup has 9 modules'          ($cleanup.Count -eq 9)
Check 'cleanup includes directory_cleaner' ($cleanup -contains 'directory_cleaner')
Check 'cleanup includes bloatware_remove'  ($cleanup -contains 'bloatware_remove')
Check 'cleanup includes destroy_history'   ($cleanup -contains 'destroy_history')
Check 'cleanup includes destroy_ssid'      ($cleanup -contains 'destroy_ssid')
Check 'cleanup includes local_user_delete' ($cleanup -contains 'local_user_delete')

$copy = Get-CategoryModuleCompletion -CategoryId 'copy'
Check 'copy has 2 modules'             ($copy.Count -eq 2)
Check 'copy includes copyfile_config'  ($copy -contains 'copyfile_config')
Check 'copy includes robocopy_config'  ($copy -contains 'robocopy_config')

$install = Get-CategoryModuleCompletion -CategoryId 'install'
Check 'install has 12 modules'         ($install.Count -eq 12)
Check 'install includes app_config'    ($install -contains 'app_config')
Check 'install includes winget_install' ($install -contains 'winget_install')
Check 'install includes printer_driver_install' ($install -contains 'printer_driver_install')
Check 'install includes printer_register'       ($install -contains 'printer_register')

$scripting = Get-CategoryModuleCompletion -CategoryId 'scripting'
Check 'scripting has 5 modules'        ($scripting.Count -eq 5)
Check 'scripting includes script_looper' ($scripting -contains 'script_looper')

Write-Host '--- Test-ModuleInCategory ---' -ForegroundColor Cyan
Check 'reg_hklm_config in settings'    (Test-ModuleInCategory -CategoryId 'settings' -ModuleName 'reg_hklm_config')
Check 'directory_cleaner in cleanup'    (Test-ModuleInCategory -CategoryId 'cleanup' -ModuleName 'directory_cleaner')
Check 'reg_hklm_config NOT in cleanup'  (-not (Test-ModuleInCategory -CategoryId 'cleanup' -ModuleName 'reg_hklm_config'))
Check 'directory_cleaner NOT in settings' (-not (Test-ModuleInCategory -CategoryId 'settings' -ModuleName 'directory_cleaner'))
Check 'hostname_config (excluded) NOT in any' (-not (Test-ModuleInCategory -CategoryId 'settings' -ModuleName 'hostname_config'))

Write-Host '--- GlobalConfig vocabulary (Phase 9 verbs) ---' -ForegroundColor Cyan
$vocab = Get-FabriqIosCommandVocabulary -Mode 'GlobalConfig'
Check 'vocab contains module'          ($vocab -contains 'module')
Check 'vocab contains cleanup'         ($vocab -contains 'cleanup')
Check 'vocab contains copy'            ($vocab -contains 'copy')
Check 'vocab contains install'         ($vocab -contains 'install')
Check 'vocab contains script'          ($vocab -contains 'script')

Write-Host '--- Tab completion routing per verb ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'

$r = Get-FabriqIosCompletion -Line 'cleanup ' -Position 8 -Mode 'GlobalConfig' -State $state
Check 'cleanup <space> -> cleanup mods' ($r -contains 'directory_cleaner' -and $r.Count -ge 7)
Check 'cleanup <space> excludes settings' (-not ($r -contains 'reg_hklm_config'))

$r = Get-FabriqIosCompletion -Line 'copy ' -Position 5 -Mode 'GlobalConfig' -State $state
Check 'copy <space> -> 2 candidates'    ($r.Count -eq 2)
Check 'copy <space> -> copyfile_config' ($r -contains 'copyfile_config')

$r = Get-FabriqIosCompletion -Line 'install ' -Position 8 -Mode 'GlobalConfig' -State $state
Check 'install <space> -> install mods' ($r -contains 'app_config' -and $r.Count -ge 11)

$r = Get-FabriqIosCompletion -Line 'script ' -Position 7 -Mode 'GlobalConfig' -State $state
Check 'script <space> -> scripting mods' ($r -contains 'script_looper' -and $r.Count -ge 5)

# module verb (G case: settings only)
$r = Get-FabriqIosCompletion -Line 'module ' -Position 7 -Mode 'GlobalConfig' -State $state
Check 'module <space> -> settings mods'    ($r -contains 'reg_hklm_config')
Check 'module <space> excludes cleanup'    (-not ($r -contains 'directory_cleaner'))
Check 'module <space> excludes copy'       (-not ($r -contains 'copyfile_config'))

Write-Host '--- Wrong-category rejection ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$out = Get-CapturedOutput { Enter-CategoryConfigMode -Verb 'cleanup' -Name 'reg_hklm_config' -State $state }
Check 'cleanup reg_hklm_config rejected' ($out -match "is not in the 'cleanup' category")
Check 'state stays GlobalConfig (wrong cat)' ($state.Mode -eq 'GlobalConfig')

$out = Get-CapturedOutput { Enter-CategoryConfigMode -Verb 'module' -Name 'directory_cleaner' -State $state }
Check 'module directory_cleaner rejected' ($out -match "is not in the 'settings' category")
Check 'state stays GlobalConfig (wrong cat 2)' ($state.Mode -eq 'GlobalConfig')

Write-Host '--- Mode entry per category + prompt suffix ---' -ForegroundColor Cyan

# cleanup directory_cleaner -> ModuleConfig with config-clean prompt
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'cleanup' -Name 'directory_cleaner' -State $state | Out-Null
Check 'cleanup -> ModuleConfig'         ($state.Mode -eq 'ModuleConfig')
Check 'CurrentCategoryId = cleanup'    ($state.CurrentCategoryId -eq 'cleanup')
Check 'prompt is config-clean'         ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-clean)#')

# exit clears category
$resolved = @{ Command = 'exit'; Args = @(); Error = $null }
Invoke-ModuleConfigCommand -Resolved $resolved -State $state | Out-Null
Check 'exit clears CurrentCategoryId'  ($null -eq $state.CurrentCategoryId)

# copy copyfile_config -> config-copy
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'copy' -Name 'copyfile_config' -State $state | Out-Null
Check 'copy -> ModuleConfig'           ($state.Mode -eq 'ModuleConfig')
Check 'prompt is config-copy'          ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-copy)#')

# install app_config -> config-install
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'install' -Name 'app_config' -State $state | Out-Null
Check 'install -> ModuleConfig'        ($state.Mode -eq 'ModuleConfig')
Check 'prompt is config-install'       ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-install)#')

# script script_looper -> config-script
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'script' -Name 'script_looper' -State $state | Out-Null
Check 'script -> ModuleConfig'         ($state.Mode -eq 'ModuleConfig')
Check 'prompt is config-script'        ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-script)#')

# module reg_hklm_config -> config-mod (settings)
$state = New-ShellState
$state.Mode = 'GlobalConfig'
Enter-CategoryConfigMode -Verb 'module' -Name 'reg_hklm_config' -State $state | Out-Null
Check 'module -> ModuleConfig'         ($state.Mode -eq 'ModuleConfig')
Check 'prompt is config-mod'           ((Get-FabriqIosPrompt -State $state) -eq 'fabriq(config-mod)#')

Write-Host '--- Parser integration: full path through dispatch ---' -ForegroundColor Cyan
$state = New-ShellState
$state.Mode = 'GlobalConfig'

# parser: cleanup directory_cleaner -> dispatcher -> mode entry
$tokens = ConvertTo-FabriqIosTokens 'cleanup directory_cleaner'
$resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode 'GlobalConfig'
Check 'parser cleanup -> resolved'      ($resolved.Command -eq 'cleanup')
Invoke-GlobalConfigCommand -Resolved $resolved -State $state | Out-Null
Check 'dispatcher cleanup -> ModuleConfig' ($state.Mode -eq 'ModuleConfig')
Check 'dispatcher cleanup -> CurrentCategoryId' ($state.CurrentCategoryId -eq 'cleanup')

# Abbreviation: 'cl' should resolve to 'cleanup' uniquely
$state = New-ShellState
$state.Mode = 'GlobalConfig'
$expanded = Expand-FabriqIosAbbreviation -Tokens @('cl','directory_cleaner') -Mode 'GlobalConfig'
Check 'cl -> cleanup abbreviation'      ($expanded[0] -eq 'cleanup')

# 'co' is ambiguous: copy vs ?? (no other co* verbs in GlobalConfig vocab actually)
# Actually let me check: hostname/interface/module/cleanup/copy/install/script/exit/end/help/?
# 'co' -> 'copy' (only co* word). Good.
$expanded = Expand-FabriqIosAbbreviation -Tokens @('co','copyfile_config') -Mode 'GlobalConfig'
Check 'co -> copy abbreviation'         ($expanded[0] -eq 'copy')

# 'ins' -> 'install' ('i' / 'in' are ambiguous with 'interface')
$expanded = Expand-FabriqIosAbbreviation -Tokens @('ins','app_config') -Mode 'GlobalConfig'
Check 'ins -> install abbreviation'     ($expanded[0] -eq 'install')

# 's' is ambiguous: script vs ??? actually only 'script' starts with 's' in GlobalConfig.
# Actually let me check: hostname (h), interface (i), module (m), cleanup (c),
# copy (c), install (i), script (s), exit (e), end (e), help (h), ? - so:
# - 'c' is ambiguous (cleanup vs copy)
# - 'cl' -> cleanup
# - 'co' -> copy
# - 's' -> script (unique)
$expanded = Expand-FabriqIosAbbreviation -Tokens @('s','script_looper') -Mode 'GlobalConfig'
Check 's -> script abbreviation'        ($expanded[0] -eq 'script')

Write-Host ''
Write-Host ("=== Summary: {0} passed, {1} failed ===" -f $script:Pass, $script:Fail) `
    -ForegroundColor (& { if ($script:Fail -eq 0) { 'Green' } else { 'Red' } })
exit $script:Fail
