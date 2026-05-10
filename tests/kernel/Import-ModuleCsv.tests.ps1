# ========================================
# Pester v5 unit tests for Import-ModuleCsv
# ========================================
# Function: kernel/common.ps1 :: Import-ModuleCsv
# Run    : pwsh ./dev/run_tests.ps1
#
# Pins the Segment-strict-match contract (KERNEL_API.md §1.2) and the
# observable return-shape behavior at the caller boundary. Note the
# scalar-unwrap quirk pinned in the "Caller-observable boundaries"
# Context: Import-ModuleCsv documents a $null vs @() distinction
# internally, but PowerShell collapses @() to $null at the assignment
# site, so callers cannot tell them apart.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    . "$PSScriptRoot\..\_helpers\test_csv.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'Import-ModuleCsv' {

    Context 'Segment strict match (KERNEL_API.md §1.2)' {

        It 'returns all rows when CSV has no Segment column' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName') -Rows @(
                @{ Enabled = '1'; TargetName = 'a' }
                @{ Enabled = '1'; TargetName = 'b' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -Segment 'alpha'
                @($r).Count | Should -Be 2
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'matches empty-Segment rows when Segment param is empty' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName','Segment') -Rows @(
                @{ Enabled = '1'; TargetName = 'a'; Segment = '' }
                @{ Enabled = '1'; TargetName = 'b'; Segment = 'alpha' }
                @{ Enabled = '1'; TargetName = 'c'; Segment = '' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -Segment ''
                @($r).Count | Should -Be 2
                (@($r) | ForEach-Object TargetName) | Should -Be @('a','c')
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'matches only same-value rows when Segment param is non-empty' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName','Segment') -Rows @(
                @{ Enabled = '1'; TargetName = 'a'; Segment = 'alpha' }
                @{ Enabled = '1'; TargetName = 'b'; Segment = 'beta'  }
                @{ Enabled = '1'; TargetName = 'c'; Segment = 'alpha' }
                @{ Enabled = '1'; TargetName = 'd'; Segment = ''      }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -Segment 'alpha'
                @($r).Count | Should -Be 2
                (@($r) | ForEach-Object TargetName) | Should -Be @('a','c')
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'trims whitespace on both Segment param and row Segment value' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName','Segment') -Rows @(
                @{ Enabled = '1'; TargetName = 'a'; Segment = '  alpha  ' }
                @{ Enabled = '1'; TargetName = 'b'; Segment = 'beta' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -Segment '  alpha  '
                @($r).Count | Should -Be 1
                @($r)[0].TargetName | Should -Be 'a'
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'normalizes whitespace-only Segment param to empty (matches empty rows)' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName','Segment') -Rows @(
                @{ Enabled = '1'; TargetName = 'a'; Segment = '' }
                @{ Enabled = '1'; TargetName = 'b'; Segment = 'alpha' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -Segment '   '
                @($r).Count | Should -Be 1
                @($r)[0].TargetName | Should -Be 'a'
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context '-FilterEnabled' {

        It 'returns only Enabled=1 rows when -FilterEnabled is set' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName') -Rows @(
                @{ Enabled = '1'; TargetName = 'a' }
                @{ Enabled = '0'; TargetName = 'b' }
                @{ Enabled = '1'; TargetName = 'c' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -FilterEnabled
                @($r).Count | Should -Be 2
                (@($r) | ForEach-Object TargetName) | Should -Be @('a','c')
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'returns all rows (including Enabled=0) when -FilterEnabled is not set' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName') -Rows @(
                @{ Enabled = '1'; TargetName = 'a' }
                @{ Enabled = '0'; TargetName = 'b' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv
                @($r).Count | Should -Be 2
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context '-RequiredColumns' {

        It 'returns rows when all required columns are present' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName','Description') -Rows @(
                @{ Enabled = '1'; TargetName = 'a'; Description = 'desc-a' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -RequiredColumns @('Enabled','TargetName')
                @($r).Count | Should -Be 1
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'returns $null when a required column is missing' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName') -Rows @(
                @{ Enabled = '1'; TargetName = 'a' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -RequiredColumns @('Enabled','TargetName','MissingCol')
                $r | Should -BeNullOrEmpty
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context 'Caller-observable boundaries (PowerShell scalar-unwrap quirk)' {

        # Import-ModuleCsv internally distinguishes:
        #   $null  : load error / file empty / RequiredColumns failure
        #   @()    : filter eliminated all rows (Enabled / Segment)
        # However, `return @()` from a function gets unwrapped to $null
        # when the caller assigns to a scalar variable. So all of the
        # cases below are observed as $null at the call site, and
        # standard module template's `if ($items.Count -eq 0)` branch is
        # effectively unreachable.

        It 'collapses to $null when -FilterEnabled eliminates all rows' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName') -Rows @(
                @{ Enabled = '0'; TargetName = 'a' }
                @{ Enabled = '0'; TargetName = 'b' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -FilterEnabled
                $null -eq $r | Should -BeTrue
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'collapses to $null when Segment filter eliminates all rows' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName','Segment') -Rows @(
                @{ Enabled = '1'; TargetName = 'a'; Segment = 'beta' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -Segment 'alpha'
                $null -eq $r | Should -BeTrue
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'returns $null when the file does not exist' {
            $missing = Join-Path $env:TEMP ("fabriq-test-missing-{0}.csv" -f ([guid]::NewGuid().ToString('N')))
            $r = Import-ModuleCsv -Path $missing
            $r | Should -BeNullOrEmpty
        }

        It 'returns $null when the file is empty (header only, no data rows)' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName') -Rows @()
            try {
                $r = Import-ModuleCsv -Path $csv
                $r | Should -BeNullOrEmpty
            } finally { Remove-TestProfileCsv $csv }
        }
    }
}
