# Dispatcher for Module Configuration mode.
#
# Cisco IOS semantics: each `set` / `add` immediately invokes the
# bound module with a single ephemeral row built from the column/
# value pairs. There is no commit / queue - the command IS the
# transaction, mirroring how `(config-if)# ip address ...` applies
# at the moment of typing in real IOS.

function Invoke-ModuleConfigCommand {
    param(
        [hashtable]$Resolved,
        [hashtable]$State
    )
    switch ($Resolved.Command) {
        'set' { Invoke-ModuleEphemeralRun -PairArgs $Resolved.Args -State $State }
        'add' { Invoke-ModuleEphemeralRun -PairArgs $Resolved.Args -State $State }
        'show' { Show-ModuleConfigSchema -State $State }
        'help' { Show-FabriqIosHelp -Mode $State.Mode }
        '?'    { Show-FabriqIosHelp -Mode $State.Mode }
        'exit' {
            $State.ConfigModuleName   = $null
            $State.ConfigModuleSchema = $null
            $State.CurrentCategoryId  = $null
            Set-ShellMode -State $State -NewMode 'GlobalConfig'
        }
        'end' {
            $State.ConfigModuleName   = $null
            $State.ConfigModuleSchema = $null
            $State.CurrentCategoryId  = $null
            Set-ShellMode -State $State -NewMode 'PrivilegedExec'
        }
        default {
            Write-Host ("% Unknown ModuleConfig command: {0}" -f $Resolved.Command) -ForegroundColor Red
        }
    }
}
