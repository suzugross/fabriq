# ========================================
# Pester v5 unit tests for Get-FlexModuleStateMap
# ========================================
# Function: apps/fabriq_operator/lib/flex_dashboard.ps1 :: Get-FlexModuleStateMap
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins which recorded results the FlexProfile dashboard shows for a row
# (TM t-0111). A result belongs to (SessionID, Profile, Order): Order is
# local to a profile, so a result recorded for ANOTHER profile in the same
# session must never paint a row of the profile on screen. The real bug:
# profile A ran "Hostname Change" at Order 10 (Success); opening profile B
# showed its Order 10 "App Install" row as Success although it never ran.
# Two independent paths produced it and both are pinned here:
#   - the history.csv overlay (matched on Order only)
#   - the $script:LastBatchResults overlay (passed in unscoped)
#
# Only the pure state resolver is unit-tested; the WinForms dashboard
# (Show-FlexDashboard) is verified manually / on the VM rig.
# History rows are hand-written with explicit timestamps (the writer has
# second resolution, so "latest wins" cannot be tested through it).
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
    . (Join-Path $script:RepoRoot 'apps\fabriq_operator\lib\flex_dashboard.ps1')

    function New-ProfileRow {
        param([Parameter(Mandatory)][int]$Order, [Parameter(Mandatory)][string]$MenuName)
        [PSCustomObject]@{ Order = $Order; MenuName = $MenuName; Category = 'Test'; RelativePath = '' }
    }

    # Writes history.csv with the current 14-column header. Each row is the
    # part AFTER the PCName column:
    #   ModuleName,Status,Message,SessionID,Verified,Order,Profile
    function Set-TestHistory {
        param([string[]]$Rows = @())
        $lines = @('Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID,Verified,Order,Profile')
        $minute = 0
        foreach ($r in $Rows) {
            $f = $r.Split(',')
            $ts = '2026-10-09 10:{0:D2}:00' -f $minute
            $minute++
            $lines += ('{0},K-1,PC1,{1},Test,{2},{3},u,w,m,{4},{5},{6},{7}' -f $ts, $f[0], $f[1], $f[2], $f[3], $f[4], $f[5], $f[6])
        }
        $lines | Set-Content -Path $script:tmpHistory -Encoding UTF8
    }

    # One entry of $script:LastBatchResults as Invoke-BatchExecution publishes it.
    function New-BatchEntry {
        param([string]$ProfileName, [int]$Order, [string]$MenuName, [string]$Status = 'Success',
              $Verified = $null, [string]$Message = '')
        @{ ProfileName = $ProfileName; Order = $Order; MenuName = $MenuName
           Status = $Status; Verified = $Verified; Message = $Message }
    }

    $script:ProfileA = @(
        (New-ProfileRow -Order 10 -MenuName 'Hostname Change'),
        (New-ProfileRow -Order 20 -MenuName 'Network Config')
    )
    $script:ProfileB = @(
        (New-ProfileRow -Order 10 -MenuName 'App Install'),
        (New-ProfileRow -Order 20 -MenuName 'Printer Install'),
        (New-ProfileRow -Order 30 -MenuName 'Finalize')
    )
}

Describe 'Get-FlexModuleStateMap' {

    BeforeEach {
        $script:tmpHistory = Join-Path $env:TEMP `
            ("fabriq-flexmap-{0}.csv" -f ([guid]::NewGuid().ToString('N')))
        Set-FabriqTestState -HistoryPath $script:tmpHistory -SessionID 'sess-X'
        $script:savedKanri = $env:SELECTED_KANRI_NO
        $env:SELECTED_KANRI_NO = 'K-1'
    }

    AfterEach {
        Remove-Item $script:tmpHistory -Force -ErrorAction SilentlyContinue
        $env:SELECTED_KANRI_NO = $script:savedKanri
    }

    Context 'Another profile ran in the same session (real-bug pin)' {

        It 'history overlay: a same-Order row of another profile leaves this row Pending' {
            Set-TestHistory -Rows @('Hostname Change,Success,renamed to PC-001,sess-X,True,10,A')

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $map[10].Status   | Should -Be 'Pending'
            $map[10].Verified | Should -Be $null
            $map[10].Message  | Should -Be ''
            $map[20].Status   | Should -Be 'Pending'
            $map[30].Status   | Should -Be 'Pending'
        }

        It 'LastBatchResults overlay: the last batch of another profile leaves this row Pending' {
            # No host selected -> the history overlay is skipped entirely,
            # which isolates the in-process path.
            $env:SELECTED_KANRI_NO = ''
            $lastBatchOfA = @((New-BatchEntry -ProfileName 'A' -Order 10 -MenuName 'Hostname Change' `
                -Verified $true -Message 'renamed to PC-001'))

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B' -LastBatchResults $lastBatchOfA

            $map[10].Status | Should -Be 'Pending'
        }

        It 'both overlays together still leave the other profile untouched' {
            Set-TestHistory -Rows @('Hostname Change,Success,renamed to PC-001,sess-X,True,10,A')
            $lastBatchOfA = @((New-BatchEntry -ProfileName 'A' -Order 10 -MenuName 'Hostname Change' -Verified $true))

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B' -LastBatchResults $lastBatchOfA

            @($map.Values | Where-Object { $_.Status -ne 'Pending' }).Count | Should -Be 0
        }

        It 'the SAME module at the same Order in another profile does not leak either' {
            # A MenuName check could not separate these two rows.
            $twin = @((New-ProfileRow -Order 10 -MenuName 'Hostname Change'))
            Set-TestHistory -Rows @('Hostname Change,Success,ran for A,sess-X,True,10,A')
            $lastBatchOfA = @((New-BatchEntry -ProfileName 'A' -Order 10 -MenuName 'Hostname Change'))

            $map = Get-FlexModuleStateMap -Rows $twin -ProfileName 'A_variant' -LastBatchResults $lastBatchOfA

            $map[10].Status | Should -Be 'Pending'
        }

        It 'each profile keeps its own state: going back to the first profile shows its results again' {
            Set-TestHistory -Rows @(
                'Hostname Change,Success,renamed,sess-X,True,10,A',
                'App Install,Error,installer failed,sess-X,,10,B'
            )

            $mapA = Get-FlexModuleStateMap -Rows $script:ProfileA -ProfileName 'A'
            $mapB = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $mapA[10].Status   | Should -Be 'Success'
            $mapA[10].Verified | Should -BeTrue
            $mapA[20].Status   | Should -Be 'Pending'
            $mapB[10].Status   | Should -Be 'Error'
            $mapB[10].Message  | Should -Be 'installer failed'
        }

        It 'a later result of another profile does not mask this profile''s own failure' {
            Set-TestHistory -Rows @(
                'App Install,Error,installer failed,sess-X,,10,B',
                'Hostname Change,Success,renamed,sess-X,True,10,A'
            )

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $map[10].Status | Should -Be 'Error'
        }
    }

    Context 'This profile''s own results' {

        It 'applies the history row of this profile (Verified text -> bool / null)' {
            Set-TestHistory -Rows @(
                'App Install,Success,installed,sess-X,True,10,B',
                'Printer Install,Success,added,sess-X,False,20,B',
                'Finalize,Skipped,nothing to do,sess-X,,30,B'
            )

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $map[10].Status   | Should -Be 'Success'
            $map[10].Verified | Should -BeTrue
            $map[10].Message  | Should -Be 'installed'
            $map[20].Verified | Should -BeFalse
            $map[30].Status   | Should -Be 'Skipped'
            $map[30].Verified | Should -Be $null
        }

        It 'the latest row per Order wins (operator reset to Pending after a Success)' {
            Set-TestHistory -Rows @(
                'App Install,Success,installed,sess-X,True,10,B',
                'App Install,Pending,Reset by operator,sess-X,,10,B'
            )

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $map[10].Status  | Should -Be 'Pending'
            $map[10].Message | Should -Be 'Reset by operator'
        }

        It 'LastBatchResults of this profile overrides the history row' {
            Set-TestHistory -Rows @('App Install,Error,first attempt,sess-X,,10,B')
            $lastBatch = @((New-BatchEntry -ProfileName 'B' -Order 10 -MenuName 'App Install' `
                -Status 'Success' -Verified $true -Message 'second attempt'))

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B' -LastBatchResults $lastBatch

            $map[10].Status   | Should -Be 'Success'
            $map[10].Verified | Should -BeTrue
            $map[10].Message  | Should -Be 'second attempt'
        }

        It 'two rows running the same module are tracked by Order, not MenuName' {
            $siblings = @(
                (New-ProfileRow -Order 80  -MenuName 'Reg Config [seg:a]'),
                (New-ProfileRow -Order 110 -MenuName 'Reg Config [seg:a]')
            )
            Set-TestHistory -Rows @('Reg Config [seg:a],Success,applied,sess-X,True,80,B')

            $map = Get-FlexModuleStateMap -Rows $siblings -ProfileName 'B'

            $map[80].Status  | Should -Be 'Success'
            $map[110].Status | Should -Be 'Pending'
        }
    }

    Context 'Results outside this profile''s context' {

        It 'ignores ad-hoc runs (no Order, no profile) even when the MenuName matches' {
            Set-TestHistory -Rows @('App Install,Success,ad-hoc from the Modules tab,sess-X,True,,')

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $map[10].Status | Should -Be 'Pending'
        }

        It 'ignores a menu-mode batch (Order set, no profile) and its LastBatchResults' {
            Set-TestHistory -Rows @('App Install,Success,menu-mode batch,sess-X,True,10,')
            $menuBatch = @((New-BatchEntry -ProfileName '' -Order 10 -MenuName 'App Install'))

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B' -LastBatchResults $menuBatch

            $map[10].Status | Should -Be 'Pending'
        }

        It 'ignores rows of another session' {
            Set-TestHistory -Rows @('App Install,Success,previous session,sess-OLD,True,10,B')

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $map[10].Status | Should -Be 'Pending'
        }

        It 'ignores a LastBatchResults entry without a ProfileName key' {
            $env:SELECTED_KANRI_NO = ''
            $unscoped = @(@{ Order = 10; MenuName = 'App Install'; Status = 'Success'; Verified = $true; Message = '' })

            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B' -LastBatchResults $unscoped

            $map[10].Status | Should -Be 'Pending'
        }
    }

    Context 'Edge inputs' {

        It 'returns an empty map for a profile without rows' {
            $map = Get-FlexModuleStateMap -Rows @() -ProfileName 'B'
            $map.Count | Should -Be 0
        }

        It 'seeds every row as Pending when there is no history file' {
            $map = Get-FlexModuleStateMap -Rows $script:ProfileB -ProfileName 'B'

            $map.Count | Should -Be 3
            @($map.Values | Where-Object { $_.Status -ne 'Pending' }).Count | Should -Be 0
        }
    }
}
