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
            if ($Resolved.Args.Count -ge 3 -and $Resolved.Args.Count -le 6) {
                # Positional override: <ip> <mask> [<gw> [<dns1> [<dns2>]]]
                # Trailing args are optional. When all 5 are supplied
                # the operation is fully self-contained and works
                # without a bound host.
                $ip   = $Resolved.Args[1]
                $mask = $Resolved.Args[2]
                $gw   = if ($Resolved.Args.Count -ge 4) { $Resolved.Args[3] } else { '' }
                $dns1 = if ($Resolved.Args.Count -ge 5) { $Resolved.Args[4] } else { '' }
                $dns2 = if ($Resolved.Args.Count -ge 6) { $Resolved.Args[5] } else { '' }
                Invoke-IpAddressManual -Ip $ip -Mask $mask `
                                       -Gateway $gw -Dns1 $dns1 -Dns2 $dns2 `
                                       -State $State
            } else {
                Write-Host "% Usage: 'ip address <ip> <mask> [<gw> [<dns1> [<dns2>]]]'" -ForegroundColor Red
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
