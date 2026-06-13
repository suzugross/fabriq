# ========================================
# Pester v5 unit tests for Write-ExecutionHistory
# ========================================
# Function: kernel/common.ps1 :: Write-ExecutionHistory
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Focus (TM t-0029): the host PIN must never reach the execution history
# CSV / HTML checklist if a module accidentally echoes it into a result
# message. PC name / KanriNo stay in cleartext by design (the checklist
# is a deliverable); only SELECTED_PIN is hard-redacted.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'Write-ExecutionHistory PIN redaction' {

    BeforeEach {
        $script:tmpHistoryPath = Join-Path $env:TEMP `
            ("fabriq-history-{0}.csv" -f ([guid]::NewGuid().ToString('N')))
        Set-FabriqTestState -HistoryPath $script:tmpHistoryPath -SessionID 'test-session-001'
        $script:SessionInfo = $null
        $env:SELECTED_KANRI_NO   = 'K-001'
        $env:SELECTED_NEW_PCNAME = 'PC-NEW-01'
    }

    AfterEach {
        if (Test-Path $script:tmpHistoryPath) {
            Remove-Item $script:tmpHistoryPath -Force -ErrorAction SilentlyContinue
        }
        Remove-Item Env:\SELECTED_PIN -ErrorAction SilentlyContinue
    }

    It 'replaces the PIN with [REDACTED] when it appears in a message' {
        $env:SELECTED_PIN = 'PIN-7391-SECRET'
        $null = Write-ExecutionHistory -ModuleName 'M' -Category 'C' -Status 'Error' `
            -Message 'BitLocker recovery used PIN-7391-SECRET during setup'
        $content = Get-Content $script:tmpHistoryPath -Raw
        $content | Should -Not -Match 'PIN-7391-SECRET'
        $content | Should -Match '\[REDACTED\]'
    }

    It 'leaves the message untouched when no PIN is set' {
        Remove-Item Env:\SELECTED_PIN -ErrorAction SilentlyContinue
        $null = Write-ExecutionHistory -ModuleName 'M' -Category 'C' -Status 'Success' `
            -Message 'Plain message with no secret'
        $content = Get-Content $script:tmpHistoryPath -Raw
        $content | Should -Match 'Plain message with no secret'
        $content | Should -Not -Match '\[REDACTED\]'
    }

    It 'does not redact PC name / KanriNo (deliverable fields stay cleartext)' {
        $env:SELECTED_PIN = 'PIN-7391-SECRET'
        $null = Write-ExecutionHistory -ModuleName 'M' -Category 'C' -Status 'Success' `
            -Message 'Hostname applied'
        $content = Get-Content $script:tmpHistoryPath -Raw
        $content | Should -Match 'PC-NEW-01'
        $content | Should -Match 'K-001'
    }
}
