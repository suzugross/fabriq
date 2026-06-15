# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: category D (destructive but revert-safe) + B (needs a fixture).
@{
    schema   = 1
    module   = 'file_delete'
    category = 'D'
    scenarios = @(
        @{
            name      = 'delete-existing'
            script    = 'file_delete.ps1'
            context   = 'noninteractive'
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # Precondition: stage a test CSV that targets the fixture paths, then create those paths.
            fixture   = @(
                @{ type = 'stage-asset'; from = 'fixtures\delete_list.test.csv'; to = 'delete_list.csv' }
                @{ type = 'create-files'; paths = @('C:\fabriq_test\scratch\del_a.txt','C:\fabriq_test\scratch\sub\del_b.txt') }
            )
            expect    = @{ status = @('Success'); verified = $true }
            # Independent oracle: targets must be ABSENT after the run.
            oracle    = @{ type = 'file-exists'; mode = 'absent'
                           paths = @('C:\fabriq_test\scratch\del_a.txt','C:\fabriq_test\scratch\sub') }
            idempotent = @{ secondRun = $null }          # not naturally idempotent (IfNotFound=Error => 2nd run Fails on gone targets); C7 models this via re-fixture
            cleanup   = 'undo'                           # in-guest teardown restores the staged CSV + clears scratch (no manual revert needed)
            teardown  = @(
                @{ type = 'restore-asset'; path = 'delete_list.csv' }           # restore the shipped CSV (auto-backed-up on stage)
                @{ type = 'delete-files'; paths = @('C:\fabriq_test\scratch') } # remove fixture leftovers
            )
            notes     = 'Ships delete_list.csv all Enabled=0; a test CSV (fixtures\delete_list.test.csv) targeting the fixture paths is staged. oracle file-exists supports mode=present|absent.'
        }
    )
}
