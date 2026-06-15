# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'taskbar_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'taskbar_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()
            expect = @{ status = @('Success','Skipped'); verified = 'any' }   # independent oracle is authoritative; idempotent Skip is fine
            # C6: independent check that the layout XML was deployed to the Default profile
            oracle = @{ type = 'file-exists'; mode = 'present'
                        paths = @('C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml') }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'
            notes = 'Deploys taskbar LayoutModification.xml to the Default profile + sysprep source.'
        }
    )
}
