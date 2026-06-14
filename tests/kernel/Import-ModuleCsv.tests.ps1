# ========================================
# Pester v5 unit tests for Import-ModuleCsv
# ========================================
# Function: kernel/common.ps1 :: Import-ModuleCsv
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the Segment-strict-match contract (KERNEL_API.md §1.2) and the
# caller-boundary return contract: $null for a genuine load failure
# (missing file / empty file / RequiredColumns failure) vs a PRESERVED
# empty array (Count 0) for "loaded OK but all rows filtered out"
# (Enabled / Segment). Import-ModuleCsv returns ,@() specifically so the
# filtered-empty case does NOT unroll to $null at the caller's assignment
# site, which keeps the standard template's `if ($items.Count -eq 0)`
# Skip branch reachable instead of collapsing into the $null->Error path.
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

    Context 'Caller-observable return contract ($null=failure vs @()=filtered-empty)' {

        # Import-ModuleCsv distinguishes, AND preserves the distinction at
        # the caller's scalar assignment site:
        #   $null  : load error / file empty / RequiredColumns failure
        #   @()    : loaded OK but a filter (Enabled / Segment) eliminated
        #            all rows -> empty array (Count 0), NOT $null.
        # The empty case is returned as ,@() so PowerShell does not unroll
        # it to $null. This keeps the standard module template's
        # `if ($items.Count -eq 0) { Skipped }` branch reachable instead of
        # mis-firing the `if ($null -eq $items) { Error }` path.

        It 'returns an empty array (Count 0, not $null) when -FilterEnabled eliminates all rows' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName') -Rows @(
                @{ Enabled = '0'; TargetName = 'a' }
                @{ Enabled = '0'; TargetName = 'b' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -FilterEnabled
                $null -eq $r | Should -BeFalse
                @($r).Count | Should -Be 0
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'returns an empty array (Count 0, not $null) when Segment filter eliminates all rows' {
            $csv = New-TestProfileCsv -Columns @('Enabled','TargetName','Segment') -Rows @(
                @{ Enabled = '1'; TargetName = 'a'; Segment = 'beta' }
            )
            try {
                $r = Import-ModuleCsv -Path $csv -Segment 'alpha'
                $null -eq $r | Should -BeFalse
                @($r).Count | Should -Be 0
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
