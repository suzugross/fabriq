# ========================================
# Pester v5 unit tests for Resolve-ProfileModules
# ========================================
# Function: kernel/common.ps1 :: Resolve-ProfileModules
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Coverage targets (from CHANGELOG regression history):
#   - kernel 3.1.3 : per-Order tracking
#   - kernel 3.1.7 : sibling-row MenuName / Segment leak
#   - kernel 3.0.0 : retired-marker graceful degradation
#   - kernel 3.2.0 : Group column attribution
#   - kernel 2.1.0 : __ASYNC__ kill switch (async_config.json gate)
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    . "$PSScriptRoot\..\_helpers\test_csv.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')
}

Describe 'Resolve-ProfileModules' {

    BeforeAll {
        # Default: async marker is honored unless an It overrides this Mock.
        Mock Get-FabriqAsyncConfig { @{ Enabled = $true } }

        $script:HostnameModule = New-MockModule `
            -RelativePath 'standard\hostname_config\hostname_config.ps1' `
            -MenuName    'Hostname Config'
        $script:RegHklmModule = New-MockModule `
            -RelativePath 'standard\reg_hklm_config\reg_hklm_config.ps1' `
            -MenuName    'Registry HKLM'
        $script:AutoLogonModule = New-MockModule `
            -RelativePath 'standard\autologon_config\autologon_config.ps1' `
            -MenuName    'Auto Logon Config'
        $script:AllModules = @(
            $script:HostnameModule,
            $script:RegHklmModule,
            $script:AutoLogonModule
        )
    }

    Context '__AUTOPILOT__ marker' {

        It 'enables AutoPilot when Enabled=1' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__AUTOPILOT__'; Enabled = '1'; Description = '' }
                @{ Order = 20; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.AutoPilot        | Should -BeTrue
                $r.AutoPilotWaitSec | Should -Be 3
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'parses WaitSec=N from Description' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__AUTOPILOT__'; Enabled = '1'; Description = 'WaitSec=7' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.AutoPilot        | Should -BeTrue
                $r.AutoPilotWaitSec | Should -Be 7
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'does not enable AutoPilot when Enabled=0' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__AUTOPILOT__'; Enabled = '0'; Description = 'WaitSec=7' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.AutoPilot | Should -BeFalse
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'does not enable AutoPilot when Enabled=0 even with -IncludeDisabled' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__AUTOPILOT__'; Enabled = '0'; Description = 'WaitSec=7' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules -IncludeDisabled
                $r.AutoPilot | Should -BeFalse
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context '__ASYNC__ kill switch (kernel 2.1.0)' {

        It 'attaches _IsAsync=$true to subsequent modules when async enabled globally' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__ASYNC__'; Enabled = '1' }
                @{ Order = 20; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                Mock Get-FabriqAsyncConfig { @{ Enabled = $true } }
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count   | Should -Be 1
                $r.ValidModules[0]._IsAsync | Should -BeTrue
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'ignores __ASYNC__ marker when async_config.json is globally disabled' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__ASYNC__'; Enabled = '1' }
                @{ Order = 20; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                Mock Get-FabriqAsyncConfig { @{ Enabled = $false } }
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count       | Should -Be 1
                $r.ValidModules[0]._IsAsync | Should -BeFalse
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context 'DefaultAsync (kernel 3.3.0)' {

        It 'attaches _IsAsync=$true to all modules when DefaultAsync=true and no __ASYNC__ marker' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
                @{ Order = 20; ScriptPath = 'standard\reg_hklm_config\reg_hklm_config.ps1'; Enabled = '1' }
            )
            try {
                Mock Get-FabriqAsyncConfig { @{ Enabled = $true; DefaultAsync = $true } }
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count       | Should -Be 2
                $r.ValidModules[0]._IsAsync | Should -BeTrue
                $r.ValidModules[1]._IsAsync | Should -BeTrue
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'keeps _IsAsync=$false when DefaultAsync=false and no marker (legacy behavior)' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                Mock Get-FabriqAsyncConfig { @{ Enabled = $true; DefaultAsync = $false } }
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count       | Should -Be 1
                $r.ValidModules[0]._IsAsync | Should -BeFalse
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'kill switch (Enabled=$false) overrides DefaultAsync=$true' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                Mock Get-FabriqAsyncConfig { @{ Enabled = $false; DefaultAsync = $true } }
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count       | Should -Be 1
                $r.ValidModules[0]._IsAsync | Should -BeFalse
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context 'Group attribution (kernel 3.2.0)' {

        It 'attaches _Group="" when Group column is missing from CSV' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            ) -Columns @('Order','ScriptPath','Enabled','Description','Segment','ErrorMode')
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count   | Should -Be 1
                $r.ValidModules[0]._Group | Should -Be ''
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'attaches _Group=value when populated' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1'; Group = 'Network' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules[0]._Group | Should -Be 'Network'
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'trims whitespace from Group value' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1'; Group = '  Network  ' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules[0]._Group | Should -Be 'Network'
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context 'Order tracking (regression: kernel 3.1.3)' {

        It 'attaches int Order from CSV string value' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 30; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules[0].Order | Should -Be 30
                $r.ValidModules[0].Order | Should -BeOfType [int]
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'sorts ValidModules by Order ascending' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 30; ScriptPath = 'standard\reg_hklm_config\reg_hklm_config.ps1'; Enabled = '1' }
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count    | Should -Be 2
                $r.ValidModules[0].Order | Should -Be 10
                $r.ValidModules[1].Order | Should -Be 30
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context 'Sibling-row Segment isolation (regression: kernel 3.1.7)' {

        It 'does not leak Segment label between rows referencing the same module' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1'; Segment = 'alpha' }
                @{ Order = 20; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1'; Segment = 'beta'  }
                @{ Order = 30; ScriptPath = 'standard\hostname_config\hostname_config.ps1'; Enabled = '1' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count | Should -Be 3
                $r.ValidModules[0].MenuName | Should -Be 'Hostname Config [seg:alpha]'
                $r.ValidModules[1].MenuName | Should -Be 'Hostname Config [seg:beta]'
                $r.ValidModules[2].MenuName | Should -Be 'Hostname Config'
                # Source module object must not be mutated.
                $script:HostnameModule.MenuName | Should -Be 'Hostname Config'
            } finally { Remove-TestProfileCsv $csv }
        }
    }

    Context 'Special markers' {

        It 'expands __RESTART__ with _IsRestart flag and [RESTART] menu name' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__RESTART__'; Enabled = '1' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count        | Should -Be 1
                $r.ValidModules[0]._IsRestart | Should -BeTrue
                $r.ValidModules[0].MenuName  | Should -Be '[RESTART]'
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'expands __AUTO_to_admin01__ to autologon_config with _AutoLogonUser' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__AUTO_to_admin01__'; Enabled = '1' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count             | Should -Be 1
                $r.ValidModules[0]._AutoLogonUser | Should -Be 'admin01'
                $r.ValidModules[0].MenuName       | Should -Be '[AUTO:admin01] Auto Logon Config'
            } finally { Remove-TestProfileCsv $csv }
        }

        It 'collects retired __SHUTDOWN__ marker as InvalidPath (graceful degradation, kernel 3.0.0)' {
            $csv = New-TestProfileCsv -Rows @(
                @{ Order = 10; ScriptPath = '__SHUTDOWN__'; Enabled = '1' }
            )
            try {
                $r = Resolve-ProfileModules -ProfileCsvPath $csv -AllModules $script:AllModules
                $r.ValidModules.Count | Should -Be 0
                $r.InvalidPaths       | Should -Contain '__SHUTDOWN__'
            } finally { Remove-TestProfileCsv $csv }
        }
    }
}
