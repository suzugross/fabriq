# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A.
@{
    schema = 1
    module = 'local_user_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'apply'; script = 'local_user_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()   # ships 3 enabled sample users: admin_user / standard_user / remote_user
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'command'; run = "[bool](Get-LocalUser -Name admin_user -ErrorAction SilentlyContinue)"; equals = 'True' }
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'delete-localuser'; name = 'admin_user' }
                @{ type = 'delete-localuser'; name = 'standard_user' }
                @{ type = 'delete-localuser'; name = 'remote_user' }
            )
            notes = 'Creates sample local users (plaintext pw in CSV); oracle confirms admin_user exists; teardown removes all three.'
        }
        @{
            name = 'nopassword'; script = 'local_user_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            # ships no row with an empty Password -> stage one (plus a normal row in the same run)
            fixture = @(
                @{ type = 'stage-asset'; from = 'fixtures\local_user_list.nopassword.test.csv'; to = 'local_user_list.csv' }
            )
            expect = @{ status = @('Success'); verified = 'any' }
            # Independent of the module: the account exists, a (blank) password was set, and it is NOT flagged
            # "user must change password at next logon" (ADSI PasswordExpired=1 is what New-LocalUser -NoPassword leaves behind).
            oracle = @{ type = 'command'; run = '$u = Get-LocalUser -Name "fbqrig_nopw" -ErrorAction SilentlyContinue; $a = [ADSI]("WinNT://" + $env:COMPUTERNAME + "/fbqrig_nopw,user"); [bool]($u -and $u.PasswordLastSet -and ([int]$a.PasswordExpired.Value -eq 0))'; equals = 'True' }
            idempotent = @{ secondRun = $null }   # existing users are not counted as skips, so a 2nd run reports 0/0/0
            cleanup = 'undo'
            teardown = @(
                @{ type = 'delete-localuser'; name = 'fbqrig_nopw' }
                @{ type = 'delete-localuser'; name = 'fbqrig_pw' }
                @{ type = 'restore-asset'; path = 'local_user_list.csv' }
            )
            notes = 'Empty Password cell = account without a password. PasswordNeverExpires=0 on purpose: with it set, the must-change flag is masked and the oracle would pass either way.'
        }
    )
}
