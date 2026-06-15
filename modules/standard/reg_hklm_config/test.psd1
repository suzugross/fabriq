# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: category A - headless-testable as-is.
@{
    schema   = 1
    module   = 'reg_hklm_config'
    category = 'A'
    scenarios = @(
        @{
            name      = 'apply'
            script    = 'reg_hklm_config.ps1'
            context   = 'noninteractive'                 # noninteractive WinRM is sufficient
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @()                              # no precondition needed
            expect    = @{ status = @('Success'); verified = $true }
            oracle    = @{ type = 'registry-csv'; csv = 'reg_hklm_list*.csv' }
            idempotent = @{ secondRun = 'Success' }      # FORCE_OVERWRITE reapplies deterministically
            cleanup   = 'none'                           # idempotent + isolated HKLM policy keys; harmless to leave between tests (reverted at next clean-base)
            notes     = 'PoC-proven 6/6 (2026-06-15); independent registry oracle confirmed.'
        }
    )
}
