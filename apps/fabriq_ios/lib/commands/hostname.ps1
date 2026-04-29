# hostname * command implementations.

function Get-HostnameCompletionFromHostlist {
    param([hashtable]$State)
    $csvPath = Join-Path $script:FabriqRoot 'kernel\csv\hostlist.csv'
    if (-not (Test-Path $csvPath)) { return @() }

    $previousPass = $global:FabriqMasterPassphrase
    if ($State.Passphrase) { $global:FabriqMasterPassphrase = $State.Passphrase }
    try {
        $rows = @(Import-ModuleCsv -Path $csvPath)
    } finally {
        $global:FabriqMasterPassphrase = $previousPass
    }
    if (-not $rows) { return @() }
    return @($rows | ForEach-Object { $_.NewPCName } | Where-Object { $_ })
}

function Invoke-HostnameSelection {
    param(
        [string]$NewName,
        [hashtable]$State
    )
    if ($State.Mode -ne 'GlobalConfig') {
        Write-Host "% 'hostname' is only available in global configuration mode." -ForegroundColor Red
        return
    }

    $row = Find-HostlistRowByNewName -NewName $NewName -State $State
    if (-not $row) {
        Write-FabriqIosSyslog -Severity 3 -Mnemonic 'HOSTNAME' -Key 'refused' `
                              -Placeholders @{ NewName = $NewName }
        Write-Host ("  hostlist.csv has no row with NewPCName = '{0}'" -f $NewName)
        return
    }

    Set-FabriqIosHostEnvironment -Row $row

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
