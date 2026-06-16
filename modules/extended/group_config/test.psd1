# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (+fixture).
@{
    schema = 1
    module = 'group_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'group_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # ships domain-group rows (unresolvable on a WORKGROUP VM) -> create a local
            # test user and add it to a built-in group instead.
            fixture = @(
                @{ type = 'create-localuser'; name = 'FabriqRigGroupTest' }
                @{ type = 'stage-asset'; from = 'fixtures\group_list.test.csv'; to = 'group_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'any' }
            oracle = @{ type = 'state-query'
                        query = '[bool](@(Get-LocalGroupMember -Group "Backup Operators" -ErrorAction SilentlyContinue) | Where-Object { $_.Name -like "*FabriqRigGroupTest" })'
                        expect = @{ value = 'True' } }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'Remove-LocalGroupMember -Group "Backup Operators" -Member "FabriqRigGroupTest" -ErrorAction SilentlyContinue' }
                @{ type = 'expr'; run = 'Remove-LocalUser -Name "FabriqRigGroupTest" -ErrorAction SilentlyContinue' }
                @{ type = 'restore-asset'; path = 'group_list.csv' }
            )
            notes = 'Creates a local test user, adds it to Backup Operators; oracle confirms membership; teardown removes member+user and restores the shipped CSV.'
        }
    )
}
