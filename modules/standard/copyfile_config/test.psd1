# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category B.
@{
    schema = 1
    module = 'copyfile_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'apply'; script = 'copyfile_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # shipped rows are segmented (Phase1/2/3); stage a segment-neutral test CSV so all 3 copy in one run
            fixture = @(
                @{ type = 'stage-asset'; from = 'fixtures\copy_list.test.csv'; to = 'copy_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'any' }
            oracle = @{ type = 'file-exists'; mode = 'present'
                        paths = @('C:\Users\Public\Desktop\test.txt','C:\Users\Public\Desktop\test2.txt','C:\Users\Public\Desktop\test3.txt') }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'delete-files'; paths = @('C:\Users\Public\Desktop\test.txt','C:\Users\Public\Desktop\test2.txt','C:\Users\Public\Desktop\test3.txt') }
                @{ type = 'restore-asset'; path = 'copy_list.csv' }
            )
            notes = 'Copies shipped source files to Public Desktop; oracle confirms presence; teardown removes them.'
        }
        @{
            # Profile data overlay Phase 2 / W1 (plan section 12.2): the ASSET
            # FOLDER is resolved from the profile data folder. The overlay
            # source file deliberately REUSES a shipped file name (test.txt)
            # with different content, so the content check proves folder-level
            # all-or-nothing - the module-side source\test.txt was not used.
            name = 'overlay-source'; script = 'copyfile_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = ''; profileDataDir = 'C:\fabriq\profiles\_test_p2' }
            fixture = @(
                @{ type = 'expr'; run = 'New-Item -ItemType Directory -Force -Path "C:\fabriq\profiles\_test_p2\modules\copyfile_config\source" | Out-Null; [System.IO.File]::WriteAllText("C:\fabriq\profiles\_test_p2\modules\copyfile_config\source\test.txt", "PDF-SOURCE-W1", [System.Text.Encoding]::ASCII); [System.IO.File]::WriteAllLines("C:\fabriq\profiles\_test_p2\modules\copyfile_config\copy_list.csv", @("Enabled,FileName,DestPath,Overwrite,Description,Segment", "1,test.txt,C:\fabriq_test\w1_overlay,1,W1 overlay copy,"), [System.Text.Encoding]::ASCII); if (Test-Path "C:\fabriq_test\w1_overlay") { Remove-Item "C:\fabriq_test\w1_overlay" -Recurse -Force -EA SilentlyContinue }' }
            )
            expect = @{ status = @('Success'); verified = 'any' }
            oracle = @{ type = 'command'; run = '(Get-Content "C:\fabriq_test\w1_overlay\test.txt" -Raw).Trim()'; equals = 'PDF-SOURCE-W1' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq_test\w1_overlay") { Remove-Item "C:\fabriq_test\w1_overlay" -Recurse -Force -EA SilentlyContinue }' }
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq\profiles\_test_p2") { Remove-Item "C:\fabriq\profiles\_test_p2" -Recurse -Force -EA SilentlyContinue }' }
            )
            notes = 'W1 asset folder overlay: copied content must be the PDF copy (PDF-SOURCE-W1), not the shipped source\test.txt.'
        }
    )
}
