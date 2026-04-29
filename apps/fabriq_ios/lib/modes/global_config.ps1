# Dispatcher for Global Configuration mode commands.

function Invoke-GlobalConfigCommand {
    param(
        [hashtable]$Resolved,
        [hashtable]$State
    )
    switch ($Resolved.Command) {
        'hostname' {
            if ($Resolved.Args.Count -lt 1) {
                Write-Host "% Incomplete: 'hostname <NewName>'" -ForegroundColor Red
                return
            }
            Invoke-HostnameSelection -NewName $Resolved.Args[0] -State $State
        }
        'interface' {
            if ($Resolved.Args.Count -lt 1) {
                Write-Host "% Incomplete: 'interface <alias>'" -ForegroundColor Red
                return
            }
            $alias = ($Resolved.Args -join ' ')
            Set-FabriqIosCurrentInterface -Alias $alias -State $State
        }
        'module'   { Invoke-VerbModeEntry -Verb 'module'   -Resolved $Resolved -State $State }
        'cleanup'  { Invoke-VerbModeEntry -Verb 'cleanup'  -Resolved $Resolved -State $State }
        'copy'     { Invoke-VerbModeEntry -Verb 'copy'     -Resolved $Resolved -State $State }
        'install'  { Invoke-VerbModeEntry -Verb 'install'  -Resolved $Resolved -State $State }
        'script'   { Invoke-VerbModeEntry -Verb 'script'   -Resolved $Resolved -State $State }
        'exit' { Set-ShellMode -State $State -NewMode 'PrivilegedExec' }
        'end'  { Set-ShellMode -State $State -NewMode 'PrivilegedExec' }
        'help' { Show-FabriqIosHelp -Mode $State.Mode }
        '?'    { Show-FabriqIosHelp -Mode $State.Mode }
        default {
            Write-Host ("% Unknown GlobalConfig command: {0}" -f $Resolved.Command) -ForegroundColor Red
        }
    }
}
