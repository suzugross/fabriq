# interface * command implementations.

function Get-InterfaceCompletionFromAdapters {
    # Returns InterfaceAlias values from physical, non-disabled
    # adapters. Japanese aliases (e.g. the localized "Ethernet") are intentionally
    # preserved.
    $adapters = @()
    try {
        $adapters = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
                      Where-Object { $_.Status -ne 'Disabled' })
    } catch {
        return @()
    }
    return @($adapters | Select-Object -ExpandProperty InterfaceAlias)
}

function Set-FabriqIosCurrentInterface {
    param(
        [string]$Alias,
        [hashtable]$State
    )
    $State.CurrentInterface = $Alias
    Set-ShellMode -State $State -NewMode 'InterfaceConfig'
    Write-FabriqIosSyslog -Severity 6 -Mnemonic 'INTERFACE' -Key 'opened' `
                          -Placeholders @{ Alias = $Alias }
}
