# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'browser_addon_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'browser_addon_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            # forcelist present on Chrome or Edge after apply
            oracle = @{ type = 'command'
                        run = "[bool]((Get-Item 'HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist' -EA SilentlyContinue).Property) -or [bool]((Get-Item 'HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist' -EA SilentlyContinue).Property)"
                        equals = 'True' }
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'none'
            notes = 'Force-install extension policy registry. Idempotent (already-registered => Skipped).'
        }
    )
}
