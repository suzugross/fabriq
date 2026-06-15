# Test descriptor (C2 contract, schema 1). Non-shipping. Category B (wrapper).
@{
    schema = 1
    module = 'script_looper'
    category = 'B'
    scenarios = @(
        @{
            name = 'run'; script = 'script_looper.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # wrapper: runs target scripts from looper_list.csv with retry/loop. Skipped/Partial if targets absent.
            expect = @{ status = @('Success','Skipped','Partial'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # testability = that of its wrapped targets; a stub target script would make it deterministic
            idempotent = @{ secondRun = $null }
            cleanup = 'none'
            notes = 'Meta/wrapper module. Often paired with azure_ad_join_check (retry-until). Deterministic only with a controlled stub target.'
        }
    )
}
