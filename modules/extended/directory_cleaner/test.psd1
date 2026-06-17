# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: category D (destructive but revert-safe) + B (needs a fixture).
@{
    schema   = 1
    module   = 'directory_cleaner'
    category = 'D'
    scenarios = @(
        @{
            name      = 'clean-both-modes'
            script    = 'directory_cleaner.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # Precondition: stage a test CSV (directory + contents rows targeting the
            # fixture), then create the scratch tree those rows delete.
            fixture   = @(
                @{ type = 'stage-asset'; from = 'fixtures\clean_list.test.csv'; to = 'clean_list.csv' }
                @{ type = 'create-files'; paths = @(
                    'C:\fabriq_test\cleaner_scratch\dir_a\seed.txt',
                    'C:\fabriq_test\cleaner_scratch\dir_b\inside.txt') }
            )
            expect    = @{ status = @('Success'); verified = $true }
            # Independent oracle: directory-mode target fully gone AND the file inside
            # the contents-mode folder removed. (dir_b itself is retained but empty.)
            oracle    = @{ type = 'file-exists'; mode = 'absent'
                           paths = @(
                               'C:\fabriq_test\cleaner_scratch\dir_a',
                               'C:\fabriq_test\cleaner_scratch\dir_b\inside.txt') }
            idempotent = @{ secondRun = 'Skipped' }      # 2nd run: dir_a absent + dir_b empty -> both Skip
            cleanup   = 'undo'
            teardown  = @(
                @{ type = 'restore-asset'; path = 'clean_list.csv' }              # restore shipped CSV (auto-backed-up on stage)
                @{ type = 'delete-files'; paths = @('C:\fabriq_test\cleaner_scratch') }
            )
            notes     = 'Ships clean_list.csv all Enabled=0; a test CSV (fixtures\clean_list.test.csv) with a directory-mode and a contents-mode row targeting the scratch tree is staged. Verifies both modes via absence oracle.'
        }
    )
}
