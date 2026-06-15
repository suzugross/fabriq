# Test descriptor (C2 contract, schema 1). Non-shipping. Category B.
@{
    schema = 1
    module = 'app_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'install'; script = 'app_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # needs installer binaries under file\ (absent on a bare VM => Skipped)
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # presence-level only; per-installer behaviour varies
            idempotent = @{ secondRun = $null }
            cleanup = 'snapshot'   # installs apps if installers present -> revert after
            notes = 'Generic app installer. Skipped when no installers staged under file\.'
        }
    )
}
