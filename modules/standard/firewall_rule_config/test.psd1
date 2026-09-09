# Test descriptor (C2 contract, schema 1). Non-shipping test metadata (Deploy-excluded). Category A (export).
@{
    schema = 1
    module = 'firewall_rule_config'
    category = 'A'
    scenarios = @(
        @{
            name = 'export'; script = 'firewall_rule_export.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = '' }
            fixture = @()   # ships one Enabled=1 Export row (import rows are ack-gated -> Skipped)
            expect = @{ status = @('Success','Skipped'); verified = 'any' }   # independent oracle is authoritative; idempotent Skip is fine
            # C6: independent check that a policy.wfw snapshot was written under backup\<timestamp>\
            oracle = @{ type = 'command'
                        run = '[bool](Get-ChildItem "C:\fabriq\modules\standard\firewall_rule_config\backup" -Recurse -Filter policy.wfw -ErrorAction SilentlyContinue)'
                        equals = 'True' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'none'   # export appends an Import row to its CSV; -SyncRepo restores the shipped CSV each run
            notes = 'Exports the firewall policy snapshot (.wfw). The import variant is ack-gated and Skips by default (would be D / WinRM-risky).'
        }
        @{
            # Profile data overlay Phase 3 / -ForWrite (plan section 13). A second
            # write shape: the snapshot goes to backup\<timestamp>\policy.wfw and
            # the export also appends an Import row to the CSV it read. The row's
            # SourcePath is relativized against the SAME resolved backup root the
            # import script anchors relative paths under, so the pair still lines up.
            name = 'overlay-export'; script = 'firewall_rule_export.ps1'
            context = 'noninteractive'; winrmSafe = $true; reboot = $false; secrets = $false
            envelope = @{ autopilot = $true; selected = @{}; passphrase = ''; segment = ''; profileDataDir = 'C:\fabriq\profiles\_test_p3' }
            fixture = @(
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq\profiles\_test_p3") { Remove-Item "C:\fabriq\profiles\_test_p3" -Recurse -Force -EA SilentlyContinue }; $b = Join-Path $ModuleDirVM "backup"; if (Test-Path $b) { Get-ChildItem $b -Directory -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue }' }
            )
            expect = @{ status = @('Success','Skipped'); verified = 'any' }
            oracle = @{ type = 'command'
                        run = '$p = [bool](Get-ChildItem "C:\fabriq\profiles\_test_p3\modules\firewall_rule_config\backup" -Recurse -Filter policy.wfw -EA SilentlyContinue); $m = [bool](Get-ChildItem (Join-Path $ModuleDirVM "backup") -Recurse -Filter policy.wfw -EA SilentlyContinue); "$p/$m"'
                        equals = 'True/False' }
            idempotent = @{ secondRun = 'Success' }
            cleanup = 'undo'
            teardown = @(
                @{ type = 'expr'; run = 'if (Test-Path "C:\fabriq\profiles\_test_p3") { Remove-Item "C:\fabriq\profiles\_test_p3" -Recurse -Force -EA SilentlyContinue }' }
            )
            notes = 'P3 write overlay: policy.wfw lands under the PDF backup root and NOT under the module backup\. The CSV auto-registration still targets the CSV it read (module dir on fallback) - self-consistent because SourcePath is relative to the resolved backup root.'
        }
    )
}
