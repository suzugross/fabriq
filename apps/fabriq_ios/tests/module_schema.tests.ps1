# Pester tests for Get-ModuleCsvSchema's CSV resolution, including the
# non-*_list.csv fallback (TM t-0035). These read the real committed
# module CSVs (deterministic) to validate end-to-end resolution.
# Run: Invoke-Pester apps\fabriq_ios\tests\module_schema.tests.ps1

BeforeAll {
    $script:FabriqIosRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:FabriqRoot    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
    . (Join-Path $script:FabriqIosRoot 'lib\commands\categories.ps1')
    . (Join-Path $script:FabriqIosRoot 'lib\commands\module.ps1')
}

Describe 'Get-ModuleCsvSchema - non-*_list.csv fallback' {
    It 'resolves domain_join via its domain.csv (no *_list.csv present)' {
        $schema = Get-ModuleCsvSchema -Name 'domain_join' -CategoryId 'settings'
        $schema             | Should -Not -BeNullOrEmpty
        $schema.CsvFileName | Should -Be 'domain.csv'
        $schema.Columns     | Should -Contain 'domain'
        $schema.Columns     | Should -Contain 'user'
        $schema.Columns     | Should -Contain 'pass'
        $schema.Columns     | Should -Contain 'dns'
    }

    It 'still resolves a *_list.csv module via the primary glob (no regression)' {
        # autologon_config ships autologon_list.csv + autologon_config.ps1;
        # the fallback must not be reached when a *_list.csv exists.
        $schema = Get-ModuleCsvSchema -Name 'autologon_config' -CategoryId 'settings'
        $schema             | Should -Not -BeNullOrEmpty
        $schema.CsvFileName | Should -BeLike '*_list.csv'
    }
}
