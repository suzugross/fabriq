# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
@{
    schema = 1
    module = 'dpi_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'dpi_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # ships Enabled=0 / AUTO (interactive on 3 displays, raw Read-Host) -> stage a
            # specific HardwareID prefix (MSBDD*) + explicit ScalePercent=125 to force the
            # direct non-interactive path and a real apply.
            fixture = @(
                @{ type = 'stage-asset'; from = 'fixtures\dpi_list.test.csv'; to = 'dpi_list.csv' }
            )
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # module re-reads DpiValue it just wrote (HKCU + Default hive); rig surfaces -Verified
            cleanup = 'undo'
            teardown = @(
                @{ type = 'restore-asset'; path = 'dpi_list.csv' }
                @{ type = 'expr'; run = 'Get-ChildItem "HKCU:\Control Panel\Desktop\PerMonitorSettings" -EA SilentlyContinue | Where-Object { $_.PSChildName -like "MSBDD*" } | Remove-Item -Recurse -Force -EA SilentlyContinue' }
            )
            notes = 'Stages HardwareID=MSBDD* + 125% to force the non-interactive apply path; teardown removes the created HKCU PerMonitorSettings key + restores CSV. Default-hive residue is reverted at next clean-base.'
        }
    )
}
