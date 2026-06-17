# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
@{
    schema = 1
    module = 'profile_delete'
    category = 'A'
    scenarios = @(
        @{
            name = 'delete'; script = 'profile_delete.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            # Ships all rows disabled -> stage a throwaway profile folder (orphaned, no WMI record)
            # + an enabled CSV row targeting it. Module removes the folder; verify confirms it is gone.
            fixture = @(
                @{ type = 'expr'; run = '$p = "C:\Users\FabriqRigDelTest"; New-Item -ItemType Directory -Path $p -Force | Out-Null; Set-Content -Path (Join-Path $p "marker.txt") -Value "rig" -Force' }
                @{ type = 'stage-asset'; from = 'fixtures\profile_list.test.csv'; to = 'profile_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'True' }   # folder gone AND no WMI record -> Verified True
            oracle = @{ type = 'self-verified' }   # two-signal: Test-Path absent + Win32_UserProfile absence; has C:\Users path guard
            idempotent = @{ secondRun = 'Skipped' }   # folder already gone -> Skip (skip = idempotency-verified)
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'Remove-Item "C:\Users\FabriqRigDelTest" -Recurse -Force -EA SilentlyContinue' }
                @{ type = 'restore-asset'; path = 'profile_list.csv' }
            )
            notes = 'Stages a throwaway C:\Users\FabriqRigDelTest folder (orphaned, no WMI record) + an enabled CSV row; module removes it; oracle confirms -Verified (folder gone AND no Win32_UserProfile record). Teardown restores CSV + removes any residual folder.'
        }
    )
}
