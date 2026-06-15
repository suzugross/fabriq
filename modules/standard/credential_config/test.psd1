# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: oracle infeasible (on the Post-Apply Verification exclusion list).
@{
    schema   = 1
    module   = 'credential_config'
    category = 'none'
    scenarios = @(
        @{
            name      = 'add'
            script    = 'credential_config.ps1'
            context   = 'interactive'                    # cmdkey rejects SYSTEM; CredWrite needs a logon session (Win32 1312 otherwise, Wave-1 2026-06-15)
            winrmSafe = $true
            reboot    = $false
            secrets   = $false                           # passwords are PLAINTEXT in CSV (no ENC) -> still use DUMMY creds only
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @()
            expect    = @{ status = @('Success','Skipped'); verified = $null }   # ships credential_list.csv all-disabled => Skipped by default; staging a dummy test CSV would exercise the add path (Success)
            oracle    = @{ type = 'none'; reason = 'cmdkey cannot read back the password; /list false-PASSes. Run-only, no state assertion.' }
            idempotent = @{ secondRun = 'Success' }
            cleanup   = 'none'                            # ships disabled => no residue; if a dummy test CSV is staged, switch to undo with a cmdkey /delete teardown
            notes     = 'On the verification-exclusion list. Routes to the interactive session-(b) runner. Ships disabled => Skipped; stage a dummy-cred test CSV to exercise the add path.'
        }
    )
}
