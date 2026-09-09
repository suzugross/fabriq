# Test descriptor (C2 contract, schema 1). Non-shipping. SCAFFOLD (not a real kitting module).
@{
    schema = 1
    module = 'test_harness_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'simulate'; script = 'test_harness_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()
            expect = @{ status = @('Success','Partial','Skipped'); verified = 'any' }   # data-driven mix of simulated statuses
            oracle = @{ type = 'self-verified' }
            idempotent = @{ secondRun = $null }
            cleanup = 'none'
            notes = 'Framework test scaffold (data-driven Status/Verified/ErrorMode simulation). Doubles as a harness self-test.'
        }
        @{
            # Profile data overlay E2E (dev/PROFILE_DATA_OVERLAY_PLAN.md section 15.2):
            # the profile data folder carries its own CSV whose single row is
            # Behavior=skip, so a Skipped status proves the PDF copy was read
            # instead of the shipped CSV (whose default-segment row is a success).
            name = 'overlay-profile'; script = 'test_harness_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = ''; profileDataDir = 'C:\fabriq\profiles\_test_harness' }
            fixture = @(
                @{ type = 'expr'; run = 'New-Item -ItemType Directory -Force -Path "C:\fabriq\profiles\_test_harness\modules\test_harness_config" | Out-Null; [System.IO.File]::WriteAllLines("C:\fabriq\profiles\_test_harness\modules\test_harness_config\test_harness_list.csv", @("Enabled,Segment,TestName,Behavior,Verified,FailFirstN,DelaySec,Description", "1,,ov1,skip,,0,0,Overlay row from profile data folder"), [System.Text.Encoding]::ASCII)' }
            )
            expect = @{ status = @('Skipped'); verified = 'any' }
            oracle = @{ type = 'command'; run = 'Test-Path "C:\fabriq\profiles\_test_harness\modules\test_harness_config\test_harness_list.csv"'; equals = 'True' }
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq\profiles\_test_harness") { Remove-Item "C:\fabriq\profiles\_test_harness" -Recurse -Force -EA SilentlyContinue }' }
            )
            notes = 'Overlay: PDF copy (Behavior=skip) must win over the shipped CSV -> Skipped. Fixture writes the PDF CSV; teardown removes the whole test PDF folder.'
        }
        @{
            # Overlay fallback: the PDF folder exists but has no CSV for this
            # module -> the kernel warns and falls back to the shipped CSV
            # (default-segment success row) -> Success.
            name = 'overlay-fallback'; script = 'test_harness_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = ''; profileDataDir = 'C:\fabriq\profiles\_test_harness' }
            fixture = @(
                @{ type = 'expr'; run = 'New-Item -ItemType Directory -Force -Path "C:\fabriq\profiles\_test_harness\modules\test_harness_config" | Out-Null; Remove-Item "C:\fabriq\profiles\_test_harness\modules\test_harness_config\test_harness_list.csv" -Force -EA SilentlyContinue' }
            )
            expect = @{ status = @('Success'); verified = 'True' }
            oracle = @{ type = 'command'; run = 'Test-Path "C:\fabriq\profiles\_test_harness\modules\test_harness_config\test_harness_list.csv"'; equals = 'False' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq\profiles\_test_harness") { Remove-Item "C:\fabriq\profiles\_test_harness" -Recurse -Force -EA SilentlyContinue }' }
            )
            notes = 'Overlay fallback: PDF folder without the CSV -> shipped CSV is used (visible FALLBACK warning) -> Success/Verified True.'
        }
    )
}
