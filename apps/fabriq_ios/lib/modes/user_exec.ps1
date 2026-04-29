# Dispatcher for User EXEC mode commands.

function Invoke-UserExecCommand {
    param(
        [hashtable]$Resolved,
        [hashtable]$State
    )
    switch ($Resolved.Command) {
        'enable' { Invoke-Enable -State $State }
        'show'   { Invoke-ShowCommand -ArgList $Resolved.Args -State $State }
        'help'   { Show-FabriqIosHelp -Mode $State.Mode }
        '?'      { Show-FabriqIosHelp -Mode $State.Mode }
        'exit' {
            Write-FabriqIosSyslog -Severity 5 -Mnemonic 'EXIT' -Key 'seance_ends' -Placeholders @{}
            $State.ShouldExit = $true
        }
        default {
            Write-Host ("% Unknown UserExec command: {0}" -f $Resolved.Command) -ForegroundColor Red
        }
    }
}
