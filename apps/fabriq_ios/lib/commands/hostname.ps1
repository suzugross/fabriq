# hostname * command implementations.
# Phase 8: hostlist coupling removed. `(config)# hostname <NewName>`
# now takes the name as a positional ad-hoc value, sets
# SELECTED_NEW_PCNAME directly (with adhoc identity for the other
# SELECTED_* identity slots if no host is bound), and runs
# hostname_config. Subsequent `ip address` commands inherit the
# bound name for module-side display.

function Invoke-HostnameSelection {
    param(
        [string]$NewName,
        [hashtable]$State
    )
    if ($State.Mode -ne 'GlobalConfig') {
        Write-Host "% 'hostname' is only available in global configuration mode." -ForegroundColor Red
        return
    }
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        Write-Host "% Incomplete: 'hostname <NewName>'" -ForegroundColor Red
        return
    }

    # Bind the name. If no identity is in place (fresh session), seed
    # adhoc OldPCName/AdminID so the underlying module's display lines
    # are populated.
    if ([string]::IsNullOrWhiteSpace($env:SELECTED_NEW_PCNAME)) {
        $env:SELECTED_OLD_PCNAME = $env:COMPUTERNAME
        $env:SELECTED_KANRI_NO   = '0'
    }
    $env:SELECTED_NEW_PCNAME = $NewName

    $modulePath = Join-Path $script:FabriqRoot 'modules\standard\hostname_config\hostname_config.ps1'
    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) { $global:FabriqMasterPassphrase = $State.Passphrase }
    try {
        $result = Invoke-FabriqIosModule -ScriptPath $modulePath
    } finally {
        $global:FabriqMasterPassphrase = $previousPass
    }

    if (-not $result) {
        Write-FabriqIosSyslog -Severity 3 -Mnemonic 'HOSTNAME' -Key 'refused' `
                              -Placeholders @{ NewName = $NewName }
        Write-Host "  Module returned no ModuleResult."
        return
    }

    switch ($result.Status) {
        'Success' {
            Write-FabriqIosSyslog -Severity 5 -Mnemonic 'HOSTNAME' -Key 'success' `
                                  -Placeholders @{ NewName = $NewName }
            Write-FabriqIosSyslog -Severity 4 -Mnemonic 'HOSTNAME' -Key 'reboot_required' `
                                  -Placeholders @{}
        }
        'Skipped' {
            Write-Host ("% Skipped: {0}" -f $result.Message) -ForegroundColor Yellow
        }
        'Cancelled' {
            Write-Host ("% Cancelled: {0}" -f $result.Message) -ForegroundColor Yellow
        }
        default {
            Write-FabriqIosSyslog -Severity 3 -Mnemonic 'HOSTNAME' -Key 'refused' `
                                  -Placeholders @{ NewName = $NewName }
            Write-Host ("  Detail: {0}" -f $result.Message)
        }
    }
}
