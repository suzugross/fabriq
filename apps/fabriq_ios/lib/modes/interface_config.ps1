# Dispatcher for Interface Configuration mode commands.

function Invoke-InterfaceConfigCommand {
    param(
        [hashtable]$Resolved,
        [hashtable]$State
    )
    switch ($Resolved.Command) {
        'ip' {
            if ($Resolved.Args.Count -lt 1 -or $Resolved.Args[0] -ne 'address') {
                Write-Host "% Incomplete: 'ip address ...' (try 'ip address ?')" -ForegroundColor Red
                return
            }
            if ($Resolved.Args.Count -eq 2 -and $Resolved.Args[1] -eq 'from-hostlist') {
                Invoke-IpAddressFromHostlist -State $State
            } elseif ($Resolved.Args.Count -eq 3) {
                Invoke-IpAddressManual -Ip $Resolved.Args[1] -Mask $Resolved.Args[2] -State $State
            } else {
                Write-Host "% Usage: 'ip address from-hostlist' or 'ip address <ip> <mask>'" -ForegroundColor Red
            }
        }
        'exit' {
            $State.CurrentInterface = $null
            Set-ShellMode -State $State -NewMode 'GlobalConfig'
        }
        'end' {
            $State.CurrentInterface = $null
            Set-ShellMode -State $State -NewMode 'PrivilegedExec'
        }
        'help' { Show-FabriqIosHelp -Mode $State.Mode }
        '?'    { Show-FabriqIosHelp -Mode $State.Mode }
        default {
            Write-Host ("% Unknown InterfaceConfig command: {0}" -f $Resolved.Command) -ForegroundColor Red
        }
    }
}
