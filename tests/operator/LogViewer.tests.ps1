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
            [datetime]$LastWrite = ([datetime]'2026-01-01T00:00:00Z')
        )
        if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
        $lines = @()
        $lines += ('{{"ts":"2026-01-01T00:00:00.000+09:00","type":"envelope.start","module":"{0}","order":{1}}}' -f $Name, $Order)
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
        It 'returns an empty array (not $null) when no file matches the order' {
            $dir = Join-Path $TestDrive 'modules_none'
            New-TestTelemetryFile -Dir $dir -Name '0001_demo' -Order 80 -ShowLines @(
                '{"ts":"t1","type":"show.info","tag":"info","msg":"x"}'
            ) | Out-Null

            $log = Get-ModuleTelemetryLog -Order 999 -ModulesDir $dir
            # Must be a preserved empty array, NOT collapsed to $null
            # (so the caller's `.Count` stays valid). Asserted without the
            # pipeline, since piping an empty array sends zero objects.
            $log -is [array] | Should -BeTrue
            @($log).Count | Should -Be 0
        }

        It 'returns an empty array when the modules dir does not exist' {
            $log = Get-ModuleTelemetryLog -Order 1 -ModulesDir (Join-Path $TestDrive 'no_such_dir')
            @($log).Count | Should -Be 0
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
