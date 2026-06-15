# ========================================
# Pester v5 unit tests for Get-ModuleTelemetryLog
# ========================================
# Function: apps/fabriq_operator/lib/log_viewer.ps1 :: Get-ModuleTelemetryLog
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the data-access contract of the FlexProfile log viewer (t-0074):
#   - Parses only show.* events (envelope.start/end excluded) from the
#     telemetry per-module JSONL, keyed by envelope.start.order.
#   - When an Order ran multiple times, the newest run wins (selected by
#     file LastWriteTimeUtc, which is restart-proof).
#   - Missing order / missing dir -> empty array (never $null-collapse).
#   - Malformed JSONL lines are skipped, never fatal.
# Only the pure data function is unit-tested; the WinForms viewer
# (Show-ModuleLogViewer) is verified manually / on the VM rig.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'apps\fabriq_operator\lib\log_viewer.ps1')

    # Writes JSONL lines to <dir>/<name>.jsonl using UTF-8 without BOM,
    # mirroring the kernel telemetry writer (AppendAllText, no BOM).
    function script:New-TestTelemetryFile {
        param(
            [Parameter(Mandatory)][string]$Dir,
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][int]$Order,
            [string[]]$ShowLines = @(),
            [datetime]$LastWrite = ([datetime]'2026-01-01T00:00:00Z'),
            [switch]$LegacyOrder
        )
        if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
        $lines = @()
        # Mirror the production telemetry writer: `order` is always 0 and the
        # real per-entry Order lives in `profileOrder` (set from the Profile
        # context). Pass -LegacyOrder to instead emit only the plain `order`
        # field (no profileOrder), exercising the fallback path.
        if ($LegacyOrder) {
            $lines += ('{{"ts":"2026-01-01T00:00:00.000+09:00","type":"envelope.start","module":"{0}","order":{1}}}' -f $Name, $Order)
        }
        else {
            $lines += ('{{"ts":"2026-01-01T00:00:00.000+09:00","type":"envelope.start","module":"{0}","order":0,"profileOrder":{1}}}' -f $Name, $Order)
        }
        $lines += $ShowLines
        $lines += '{"ts":"2026-01-01T00:00:05.000+09:00","type":"envelope.end","status":"Error"}'
        $path = Join-Path $Dir "$Name.jsonl"
        [System.IO.File]::WriteAllLines($path, [string[]]$lines, (New-Object System.Text.UTF8Encoding($false)))
        (Get-Item $path).LastWriteTimeUtc = $LastWrite
        return $path
    }
}

Describe 'Get-ModuleTelemetryLog' {

    Context 'Basic parsing' {
        It 'returns only show.* lines with Level/Tag/Message for the matching order' {
            $dir = Join-Path $TestDrive 'modules_basic'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 80 -ShowLines @(
                '{"ts":"t1","type":"show.info","tag":"info","msg":"starting"}',
                '{"ts":"t2","type":"show.error","tag":"error","msg":"boom"}'
            ) | Out-Null

            $log = @(Get-ModuleTelemetryLog -Order 80 -ModulesDir $dir)

            $log.Count | Should -Be 2
            $log[0].Level   | Should -Be 'info'
            $log[0].Message | Should -Be 'starting'
            $log[1].Level   | Should -Be 'error'
            $log[1].Tag     | Should -Be 'error'
            $log[1].Message | Should -Be 'boom'
        }

        It 'keys on profileOrder (production: order is always 0)' {
            # Reproduces the real writer: order=0, profileOrder carries the
            # Profile Order. Matching on order would fail here.
            $dir = Join-Path $TestDrive 'modules_po'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 110 -ShowLines @(
                '{"ts":"t1","type":"show.info","tag":"info","msg":"via profileOrder"}'
            ) | Out-Null

            (@(Get-ModuleTelemetryLog -Order 110 -ModulesDir $dir)).Count | Should -Be 1
            # order==0 must NOT match a real Profile Order row.
            (@(Get-ModuleTelemetryLog -Order 0 -ModulesDir $dir)).Count | Should -Be 0
        }

        It 'falls back to the plain order field when profileOrder is absent' {
            $dir = Join-Path $TestDrive 'modules_legacy'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 70 -LegacyOrder -ShowLines @(
                '{"ts":"t1","type":"show.success","tag":"info","msg":"legacy order"}'
            ) | Out-Null

            $log = @(Get-ModuleTelemetryLog -Order 70 -ModulesDir $dir)
            $log.Count | Should -Be 1
            $log[0].Message | Should -Be 'legacy order'
        }

        It 'excludes envelope.start / envelope.end events' {
            $dir = Join-Path $TestDrive 'modules_excl'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 10 -ShowLines @(
                '{"ts":"t1","type":"show.success","tag":"verifyPass","msg":"ok"}'
            ) | Out-Null

            $log = @(Get-ModuleTelemetryLog -Order 10 -ModulesDir $dir)
            $log.Count | Should -Be 1
            $log[0].Level | Should -Be 'success'
        }
    }

    Context 'Newest run wins (restart-proof)' {
        It 'returns the most recently written file when two share the same order' {
            $dir = Join-Path $TestDrive 'modules_dup'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 50 -ShowLines @(
                '{"ts":"t1","type":"show.info","tag":"info","msg":"OLD run"}'
            ) -LastWrite ([datetime]'2026-01-01T00:00:00Z') | Out-Null
            # Post-restart re-run: sequence resets (0001 again) but the file
            # is written later in wall-clock time.
            New-TestTelemetryFile -Dir $dir -Name '0001_demo_rerun' -Order 50 -ShowLines @(
                '{"ts":"t9","type":"show.success","tag":"info","msg":"NEW run"}'
            ) -LastWrite ([datetime]'2026-06-15T00:00:00Z') | Out-Null

            $log = @(Get-ModuleTelemetryLog -Order 50 -ModulesDir $dir)
            $log.Count | Should -Be 1
            $log[0].Message | Should -Be 'NEW run'
        }
    }

    Context 'Empty / robustness contracts' {
        It 'yields zero lines under @() when no file matches the order' {
            $dir = Join-Path $TestDrive 'modules_none'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 80 -ShowLines @(
                '{"ts":"t1","type":"show.info","tag":"info","msg":"x"}'
            ) | Out-Null

            # Canonical caller idiom is @(call); empty must collapse to 0,
            # not to a single blank element (the ,@() nesting footgun).
            (@(Get-ModuleTelemetryLog -Order 999 -ModulesDir $dir)).Count | Should -Be 0
        }

        It 'yields zero lines when the modules dir does not exist' {
            (@(Get-ModuleTelemetryLog -Order 1 -ModulesDir (Join-Path $TestDrive 'no_such_dir'))).Count | Should -Be 0
        }

        It 'skips malformed JSONL lines without throwing' {
            $dir = Join-Path $TestDrive 'modules_bad'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 20 -ShowLines @(
                '{"ts":"t1","type":"show.info","tag":"info","msg":"good1"}',
                'this is not json at all',
                '{"ts":"t2","type":"show.error","tag":"error","msg":"good2"}'
            ) | Out-Null

            $log = @(Get-ModuleTelemetryLog -Order 20 -ModulesDir $dir)
            $log.Count | Should -Be 2
            $log[0].Message | Should -Be 'good1'
            $log[1].Message | Should -Be 'good2'
        }
    }
}
