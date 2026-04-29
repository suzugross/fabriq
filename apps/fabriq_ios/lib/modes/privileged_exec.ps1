# Dispatcher for Privileged EXEC mode commands.

function Invoke-PrivilegedExecCommand {
    param(
        [hashtable]$Resolved,
        [hashtable]$State
    )
    switch ($Resolved.Command) {
        'show' { Invoke-ShowCommand -ArgList $Resolved.Args -State $State }
        'configure' {
            if ($Resolved.Args.Count -lt 1 -or $Resolved.Args[0] -ne 'terminal') {
                Write-Host "% Incomplete: 'configure terminal' (or 'conf t')" -ForegroundColor Red
                return
            }
            Set-ShellMode -State $State -NewMode 'GlobalConfig'
            Write-Host 'Enter configuration commands, one per line. End with CNTL/Z.'
        }
        'reload' {
            Write-FabriqIosSyslog -Severity 4 -Mnemonic 'RESTART' -Key 'reload_dream' -Placeholders @{}
            Write-Host '% (no actual reboot occurs; Fabriq IOS reload is performative)'
        }
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
