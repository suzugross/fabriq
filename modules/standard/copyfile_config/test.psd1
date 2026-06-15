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
    )
}
