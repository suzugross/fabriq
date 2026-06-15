# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'builtin_admin_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'builtin_admin_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'command'; run = "(Get-LocalUser -Name Administrator).Enabled.ToString()"; equals = 'True' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'Disable-LocalUser -Name Administrator -ErrorAction SilentlyContinue' }
            )
            notes = 'Enables/configures the built-in Administrator (plaintext pw in CSV, dummy). Teardown disables it again.'
        }
    )
}
