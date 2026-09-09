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
        @{
            # Profile data overlay Phase 2 / W3 (plan section 12.4): the
            # cross-module WRITE target (..\sysprep_config\source) is resolved
            # through the overlay, so taskbar_config writes into the PDF and the
            # module-side sysprep source stays untouched. The fixture creates the
            # PDF folder only (an empty folder is enough to opt in) and clears
            # both possible outputs first.
            name = 'overlay-sysprep-source'; script = 'taskbar_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = ''; profileDataDir = 'C:\fabriq\profiles\_test_p2' }
            fixture = @(
                @{ type = 'expr'; run = 'New-Item -ItemType Directory -Force -Path "C:\fabriq\profiles\_test_p2\modules\sysprep_config\source" | Out-Null; Remove-Item (Join-Path (Split-Path $ModuleDirVM -Parent) "sysprep_config\source\LayoutModification.xml") -Force -EA SilentlyContinue; Remove-Item "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml" -Force -EA SilentlyContinue' }
            )
            expect = @{ status = @('Success'); verified = 'True' }
            oracle = @{ type = 'command'; run = '$p = Test-Path "C:\fabriq\profiles\_test_p2\modules\sysprep_config\source\LayoutModification.xml"; $m = Test-Path (Join-Path (Split-Path $ModuleDirVM -Parent) "sysprep_config\source\LayoutModification.xml"); "$p/$m"'; equals = 'True/False' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'Remove-Item "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml" -Force -EA SilentlyContinue' }
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq\profiles\_test_p2") { Remove-Item "C:\fabriq\profiles\_test_p2" -Recurse -Force -EA SilentlyContinue }' }
            )
            notes = 'W3 cross-module write: LayoutModification.xml lands in the PDF sysprep source and NOT in the module dir.'
        }
    )
}
