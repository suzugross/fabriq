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
            fixture = @()   # ships 3 enabled rows (Explorer AppId, Edge AppId, Chrome LinkPath) -> real 3-pin deploy
            expect = @{ status = @('Success'); verified = 'True' }   # pinned-app count (3) == items -> Verified True
            # C6 independent oracle: the layout XML was deployed to the Default profile
            oracle = @{ type = 'file-exists'; mode = 'present'
                        paths = @('C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml') }
            idempotent = @{ secondRun = 'Success' }   # always regenerates/overwrites -> Success again
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'Remove-Item "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml" -Force -EA SilentlyContinue' }
                @{ type = 'expr'; run = 'Remove-Item (Join-Path (Split-Path $ModuleDirVM -Parent) "sysprep_config\source\LayoutModification.xml") -Force -EA SilentlyContinue' }
            )
            notes = 'Ships 3 enabled rows (2 DesktopApplicationID + 1 DesktopApplicationLinkPath) -> deploys LayoutModification.xml (3 pins) to the Default profile + sysprep source. Independent file-exists oracle; -Verified asserts pinned-app count == items. Teardown removes both artifacts.'
        }
    )
}
