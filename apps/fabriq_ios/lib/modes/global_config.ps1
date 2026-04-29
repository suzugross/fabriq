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
        'module' {
            if ($Resolved.Args.Count -lt 1) {
                Write-Host "% Incomplete: 'module <name>' (try 'module ?')" -ForegroundColor Red
                return
            }
            # Phase 7 redesign: enter ModuleConfig mode rather than
            # immediately executing. Run happens via `set` / `add`
            # inside (config-mod)#. The Phase 6 immediate-run path is
            # intentionally retired.
            Enter-ModuleConfigMode -Name $Resolved.Args[0] -State $State
        }
        'exit' { Set-ShellMode -State $State -NewMode 'PrivilegedExec' }
        'end'  { Set-ShellMode -State $State -NewMode 'PrivilegedExec' }
        'help' { Show-FabriqIosHelp -Mode $State.Mode }
        '?'    { Show-FabriqIosHelp -Mode $State.Mode }
        default {
            Write-Host ("% Unknown GlobalConfig command: {0}" -f $Resolved.Command) -ForegroundColor Red
        }
    }
}
