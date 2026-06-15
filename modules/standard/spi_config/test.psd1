# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded).
# Archetype: interactive context required (user32 SPI fails in Session-0/noninteractive WinRM).
@{
    schema   = 1
    module   = 'spi_config'
    category = 'A-interactive'
    scenarios = @(
        @{
            name      = 'apply'
            script    = 'spi_config.ps1'
            context   = 'interactive'                    # PROVEN: SystemParametersInfo returns false headless (Wave-1 2026-06-15)
            winrmSafe = $true
            reboot    = $false
            secrets   = $false
            envelope  = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture   = @()
            expect    = @{ status = @('Success','Skipped'); verified = 'any' }   # idempotent: Success on first apply, Skipped ("already match") on an already-applied baseline (verified true|null accordingly)
            # SPI GET read-back also needs the interactive session. For now trust the module's own SPI GET.
            oracle    = @{ type = 'self-verified' }       # TODO(C6): AI-synthesize a state-query SPI GET oracle, run in context (b)
            idempotent = @{ secondRun = 'Skipped' }
            cleanup   = 'none'                            # idempotent UI settings; harmless to leave (re-apply converges)
            notes     = 'Routes to the interactive session-(b) runner; headless WinRM => SystemParametersInfo fails. Idempotent => accepts Success or Skipped.'
        }
    )
}
