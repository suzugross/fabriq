# Test descriptor (C2 contract, schema 1). Non-shipping. Category B.
@{
    schema = 1
    module = 'cert_config'
    category = 'B'
    scenarios = @(
        @{
            name = 'import'; script = 'cert_config.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = '' }
            fixture = @()   # needs .pfx/.cer under certs\ (shipped sample or staged); no live CA needed
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'self-verified' }   # C6 can synthesize a Cert:\ thumbprint oracle once a known test cert is staged
            idempotent = @{ secondRun = 'Skipped' }
            cleanup = 'snapshot'   # imports into LocalMachine cert store -> revert after
            notes = 'Cert store import. Skipped if certs\ empty.'
        }
    )
}
