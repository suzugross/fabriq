# ========================================
# Pester v5 unit tests for New-ModuleResult
# ========================================
# Function: kernel/common.ps1 :: New-ModuleResult
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the foundation result contract (KERNEL_API.md §1.3 / §5) that every
# standard module returns through, either directly (Cancelled paths,
# explicit Status returns) or indirectly via New-BatchResult. A regression
# in the shape (_IsModuleResult marker / Timestamp / Verified tri-state)
# would silently break execution_history Status / Verified columns and
# the FlexProfile dashboard's Status / Verified badges across the entire
# fleet. The Phase 1b suite already pins New-BatchResult; this file pins
# the underlying primitive directly.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'New-ModuleResult' {

    Context 'Status ValidateSet (KERNEL_API.md §1.3)' {

        It 'accepts <_>' -ForEach @('Success','Error','Cancelled','Skipped','Partial') {
            $r = New-ModuleResult -Status $_
            $r.Status | Should -Be $_
        }

        It 'rejects an unknown Status value' {
            # ValidateSet rejection surfaces as a parameter-binding error
            # at call time, before the function body runs.
            { New-ModuleResult -Status 'Bogus' } | Should -Throw
        }
    }

    Context 'Defaults' {

        It 'defaults Message to empty string' {
            $r = New-ModuleResult -Status 'Success'
            $r.Message | Should -Be ''
        }

        It 'defaults Details to an empty array' {
            $r = New-ModuleResult -Status 'Success'
            ,$r.Details         | Should -BeOfType [System.Object[]]
            $r.Details.Count    | Should -Be 0
        }

        It 'defaults Verified to $null (no verification performed)' {
            $r = New-ModuleResult -Status 'Success'
            $r.Verified | Should -BeNullOrEmpty
            # BeNullOrEmpty also matches @(); pin the actual $null identity:
            $null -eq $r.Verified | Should -BeTrue
        }
    }

    Context 'ModuleResult shape contract (KERNEL_API.md §5)' {

        It 'sets _IsModuleResult marker to $true' {
            $r = New-ModuleResult -Status 'Success'
            $r._IsModuleResult | Should -BeTrue
        }

        It 'populates Timestamp with a [datetime] value' {
            $r = New-ModuleResult -Status 'Success'
            $r.Timestamp | Should -BeOfType [datetime]
        }

        It 'exposes exactly the 6 documented fields (no extras)' {
            $r = New-ModuleResult -Status 'Success' -Message 'm' -Details @(1) -Verified $true
            $names = $r.PSObject.Properties.Name | Sort-Object
            $expected = @('_IsModuleResult','Details','Message','Status','Timestamp','Verified') |
                Sort-Object
            # Compare-Object returns nothing when the sets are equal; any
            # extra/missing field would surface as a SideIndicator row.
            (Compare-Object $names $expected) | Should -BeNullOrEmpty
        }
    }

    Context 'Verified pass-through (Nullable[bool])' {

        It 'propagates -Verified $true' {
            $r = New-ModuleResult -Status 'Success' -Verified $true
            $r.Verified | Should -BeTrue
        }

        It 'propagates -Verified $false even for a Success Status' {
            $r = New-ModuleResult -Status 'Success' -Verified $false
            $r.Status   | Should -Be 'Success'
            $r.Verified | Should -BeFalse
        }

        It 'preserves an explicit -Verified $null (sysprep-style un-verifiable case)' {
            $r = New-ModuleResult -Status 'Success' -Verified $null
            $null -eq $r.Verified | Should -BeTrue
        }
    }

    Context 'Details pass-through' {

        It 'preserves an array of PSCustomObject details verbatim' {
            $details = @(
                [PSCustomObject]@{ Name = 'a'; Result = 'OK' }
                [PSCustomObject]@{ Name = 'b'; Result = 'Skip' }
            )
            $r = New-ModuleResult -Status 'Partial' -Details $details
            @($r.Details).Count    | Should -Be 2
            $r.Details[0].Name     | Should -Be 'a'
            $r.Details[1].Result   | Should -Be 'Skip'
        }

        It 'preserves a string array verbatim' {
            $r = New-ModuleResult -Status 'Success' -Details @('first','second','third')
            @($r.Details).Count | Should -Be 3
            $r.Details[2]       | Should -Be 'third'
        }
    }

    Context 'Side effect: $global:_LastModuleResult fallback' {
        # New-ModuleResult stashes the result on $global:_LastModuleResult
        # as a safety net for callers that lose the pipeline value. Pin
        # the existence and the overwrite-on-each-call semantics; modules
        # rely on this for the FlexProfile single-execution path
        # (AutoConfirmMode) when the pipeline is consumed by Show-* output.

        It 'sets $global:_LastModuleResult to the returned object' {
            $r = New-ModuleResult -Status 'Success' -Message 'first'
            # Reference identity: same PSCustomObject instance.
            [object]::ReferenceEquals($r, $global:_LastModuleResult) | Should -BeTrue
        }

        It 'overwrites $global:_LastModuleResult on a subsequent call' {
            $r1 = New-ModuleResult -Status 'Success' -Message 'first'
            [object]::ReferenceEquals($r1, $global:_LastModuleResult) | Should -BeTrue
            $r2 = New-ModuleResult -Status 'Error'   -Message 'second'
            [object]::ReferenceEquals($r2, $global:_LastModuleResult) | Should -BeTrue
            [object]::ReferenceEquals($r1, $global:_LastModuleResult) | Should -BeFalse
            $global:_LastModuleResult.Message | Should -Be 'second'
        }
    }
}
