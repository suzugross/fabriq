# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
@{
    schema = 1
    module = 'display_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'display_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # ships all rows Enabled=0; AUTO would hit interactive (3 display keys, raw
            # Read-Host) -> stage a SPECIFIC HardwareID prefix (MSBDD*, unique) to force
            # the direct non-interactive path and a real apply (1280x720 != current).
            fixture = @(
                @{ type = 'stage-asset'; from = 'fixtures\display_list.test.csv'; to = 'display_list.csv' }
            )
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # module re-reads PrimSurfSize it just wrote; rig surfaces -Verified
            cleanup = 'undo'
            teardown = @(
                @{ type = 'restore-asset'; path = 'display_list.csv' }
                @{ type = 'expr'; run = '$cfg = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration" -EA SilentlyContinue | Where-Object { $_.PSChildName -like "MSBDD*" } | Select-Object -First 1; if ($cfg) { $sk = Get-ChildItem $cfg.PSPath -EA SilentlyContinue | Select-Object -First 1; if ($sk) { Set-ItemProperty -Path $sk.PSPath -Name "PrimSurfSize.cx" -Value 1024 -Type DWord -EA SilentlyContinue; Set-ItemProperty -Path $sk.PSPath -Name "PrimSurfSize.cy" -Value 768 -Type DWord -EA SilentlyContinue } }' }
            )
            notes = 'Stages a specific HardwareID (MSBDD*) + 1280x720 to force the non-interactive apply path; teardown restores CSV + PrimSurfSize to 1024x768.'
        }
    )
}
