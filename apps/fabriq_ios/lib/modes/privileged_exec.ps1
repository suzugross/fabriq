# Dispatcher for Privileged EXEC mode commands.

function Invoke-PrivilegedExecCommand {
    param(
        [hashtable]$Resolved,
        [hashtable]$State
    )
    switch ($Resolved.Command) {
        'show' { Invoke-ShowCommand -ArgList $Resolved.Args -State $State }
        'ping'       { Invoke-FabriqIosPing -State $State -ArgList $Resolved.Args }
        'traceroute' { Invoke-FabriqIosTraceroute -State $State -ArgList $Resolved.Args }
        'configure' {
            if ($Resolved.Args.Count -lt 1 -or $Resolved.Args[0] -ne 'terminal') {
                Write-Host "% Incomplete: 'configure terminal' (or 'conf t')" -ForegroundColor Red
                return
            }
            Set-ShellMode -State $State -NewMode 'GlobalConfig'
            Write-Host 'Enter configuration commands, one per line. End with CNTL/Z.'
        }
        'reload' { Invoke-FabriqIosReload -State $State }
        'disable' { Invoke-Disable -State $State }
        'help'    { Show-FabriqIosHelp -Mode $State.Mode }
        '?'       { Show-FabriqIosHelp -Mode $State.Mode }
        'exit' {
            Write-FabriqIosSyslog -Severity 5 -Mnemonic 'EXIT' -Key 'seance_ends' -Placeholders @{}
            $State.ShouldExit = $true
        }
        default {
            Write-Host ("% Unknown PrivilegedExec command: {0}" -f $Resolved.Command) -ForegroundColor Red
        }
    }
}
