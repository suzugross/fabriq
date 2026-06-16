# ========================================
# Pester v5 unit tests for Get-FabriqGateBarrier
# ========================================
# Function: kernel/common.ps1 :: Get-FabriqGateBarrier
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the __GATE__ forward-barrier contract (TM t-0073, KERNEL_API §4 /
# §8 3.6.0): the first gate whose preceding window (since the previous gate
# or profile start) contains an Error/Partial is the barrier; its Order is
# returned and the caller blocks every Order >= it. Success/Skipped/
# Cancelled/Pending (absent) never block. No gate / clean windows -> $null.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    function script:New-GateRow {
        param([Parameter(Mandatory)][int]$Order, [switch]$Gate)
        $o = [pscustomobject]@{ Order = $Order }
        if ($Gate) { $o | Add-Member -NotePropertyName '_IsGate' -NotePropertyValue $true }
        return $o
    }
    # Build a status map with guaranteed [int] keys (matches production).
    function script:New-StatusMap {
        param([hashtable]$Pairs = @{})
        $m = @{}
        foreach ($k in $Pairs.Keys) { $m[[int]$k] = $Pairs[$k] }
        return $m
    }
}

Describe 'Get-FabriqGateBarrier' {

    Context 'No barrier cases' {
        It 'returns $null when there is no gate at all (even with failures)' {
            $rows = @((New-GateRow 10), (New-GateRow 20), (New-GateRow 30))
            $map  = New-StatusMap @{ 10 = 'Error'; 20 = 'Partial' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be $null
        }

        It 'returns $null when the gate window is all Success' {
            $rows = @((New-GateRow 10), (New-GateRow 20), (New-GateRow 30 -Gate), (New-GateRow 40))
            $map  = New-StatusMap @{ 10 = 'Success'; 20 = 'Success' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be $null
        }

        It 'does not block on Skipped / Cancelled / Pending (absent) in the window' {
            $rows = @((New-GateRow 10), (New-GateRow 20), (New-GateRow 25), (New-GateRow 30 -Gate))
            $map  = New-StatusMap @{ 10 = 'Skipped'; 20 = 'Cancelled' }  # 25 absent = Pending
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be $null
        }

        It 'returns $null for a gate as the very first row (empty window)' {
            $rows = @((New-GateRow 10 -Gate), (New-GateRow 20))
            $map  = New-StatusMap @{ 20 = 'Error' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be $null
        }

        It 'returns $null when the only failure is AFTER the last gate' {
            $rows = @((New-GateRow 10), (New-GateRow 20 -Gate), (New-GateRow 30))
            $map  = New-StatusMap @{ 10 = 'Success'; 30 = 'Error' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be $null
        }

        It 'returns $null for empty rows' {
            Get-FabriqGateBarrier -Rows @() -StatusMap (New-StatusMap @{}) | Should -Be $null
        }
    }

    Context 'Barrier cases' {
        It 'returns the gate Order when its window has an Error' {
            $rows = @((New-GateRow 10), (New-GateRow 20), (New-GateRow 30 -Gate), (New-GateRow 40))
            $map  = New-StatusMap @{ 10 = 'Success'; 20 = 'Error' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be 30
        }

        It 'returns the gate Order when its window has a Partial' {
            $rows = @((New-GateRow 10), (New-GateRow 30 -Gate), (New-GateRow 40))
            $map  = New-StatusMap @{ 10 = 'Partial' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be 30
        }

        It 'returns the SECOND gate when only the second window failed' {
            $rows = @((New-GateRow 10), (New-GateRow 20 -Gate), (New-GateRow 30), (New-GateRow 40 -Gate), (New-GateRow 50))
            $map  = New-StatusMap @{ 10 = 'Success'; 30 = 'Error' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be 40
        }

        It 'returns the FIRST unsatisfied gate when multiple windows fail' {
            $rows = @((New-GateRow 10), (New-GateRow 20 -Gate), (New-GateRow 30), (New-GateRow 40 -Gate))
            $map  = New-StatusMap @{ 10 = 'Error'; 30 = 'Error' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be 20
        }

        It 'is order-insensitive to the input array (sorts by Order)' {
            $rows = @((New-GateRow 40), (New-GateRow 30 -Gate), (New-GateRow 10), (New-GateRow 20))
            $map  = New-StatusMap @{ 20 = 'Error' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be 30
        }

        It 'deterministically counts a module sharing a gate Order INSIDE that window' {
            # Malformed authoring: a module and a __GATE__ both at Order 20.
            # The secondary sort key forces non-gate before gate, so the
            # erroring module is inside the gate window => barrier 20 (never
            # $null), regardless of PS 5.1 Sort-Object instability.
            $rows = @((New-GateRow 10), (New-GateRow 20), (New-GateRow 20 -Gate), (New-GateRow 30))
            $map  = New-StatusMap @{ 20 = 'Error' }
            Get-FabriqGateBarrier -Rows $rows -StatusMap $map | Should -Be 20
        }
    }

    Context 'Post-Apply Verification (Verified=FALSE) blocks' {
        It 'blocks the gate when a window module is Success but Verified=$false' {
            $rows = @((New-GateRow 10), (New-GateRow 20 -Gate), (New-GateRow 30))
            $status = New-StatusMap @{ 10 = 'Success' }
            $verified = @{}; $verified[10] = $false
            Get-FabriqGateBarrier -Rows $rows -StatusMap $status -VerifiedMap $verified | Should -Be 20
        }

        It 'does NOT block on Verified=$true or Verified=$null (unverified)' {
            $rows = @((New-GateRow 10), (New-GateRow 15), (New-GateRow 20 -Gate))
            $status = New-StatusMap @{ 10 = 'Success'; 15 = 'Success' }
            $verified = @{}; $verified[10] = $true; $verified[15] = $null
            Get-FabriqGateBarrier -Rows $rows -StatusMap $status -VerifiedMap $verified | Should -Be $null
        }

        It 'attributes a Verified=$false to the correct window (second gate)' {
            $rows = @((New-GateRow 10), (New-GateRow 20 -Gate), (New-GateRow 30), (New-GateRow 40 -Gate))
            $status = New-StatusMap @{ 10 = 'Success'; 30 = 'Success' }
            $verified = @{}; $verified[30] = $false
            Get-FabriqGateBarrier -Rows $rows -StatusMap $status -VerifiedMap $verified | Should -Be 40
        }
    }
}
