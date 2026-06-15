# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'reg_hkcu_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'reg_hkcu_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'registry-csv'; csv = 'reg_hkcu_list*.csv' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'HKCU + Default-hive registry; oracle reads back the same session HKCU written over WinRM.'
        }
    )
}
