# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'ipv6_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'ipv6_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false   # WinRM is over IPv4, so toggling IPv6 binding is safe
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # binding state depends on CSV intent; trust module for now (C6 can synthesize Get-NetAdapterBinding oracle)
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'Get-NetAdapter | ForEach-Object { Enable-NetAdapterBinding -Name $_.Name -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue }' }
            )
            notes = 'Toggles ms_tcpip6 binding. Teardown re-enables IPv6 on all adapters.'
        }
    )
}
