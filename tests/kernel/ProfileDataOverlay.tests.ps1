# ========================================
# Pester v5 unit tests for the Profile Data Overlay (Phase 1)
# ========================================
# Functions: kernel/common.ps1 :: Resolve-ModuleDataPath /
#            Get-FabriqProfileDataDir / Set-/Clear-FabriqProfileDataContext /
#            Test-FabriqResumeDataDir / Import-ModuleCsv (overlay hook)
# Run     : powershell.exe -File ./dev/run_tests.ps1
# Plan    : dev/PROFILE_DATA_OVERLAY_PLAN.md (section 4 = the contract)
#
# Pins:
#   - no context            -> identity mapping, no display (backward compat)
#   - context + file in PDF -> profile copy, Show-Info once per batch per file
#   - context + file absent -> module path, Show-Warning once (never silent)
#   - non-module paths / already-resolved paths -> untouched
#   - ..\ cross-module references normalize before the prefix check
#   - Import-ModuleCsv reads the profile copy and still applies Segment
#   - resume guard: missing data folder recorded in the state -> refuse
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot  = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    # A real module CSV under <repo>\modules\ (read-only use).
    $script:ModuleCsv = Join-Path $script:RepoRoot 'modules\standard\taskbar_config\taskbar_list.csv'

    function New-OverlayFile {
        param([string]$Path, [string[]]$Lines)
        $dir = Split-Path $Path -Parent
        if (-not (Test-Path $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
        [System.IO.File]::WriteAllLines($Path, $Lines, [System.Text.Encoding]::ASCII)
    }

    function Remove-OverlayTemp {
        param([string]$Path)
        # Destructive-path guard (CLAUDE.md section 8): only ever delete the
        # dedicated temp folders this file created.
        $tempRoot = [System.IO.Path]::GetFullPath($env:TEMP)
        if ($Path -and $Path.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) `
            -and $Path -like '*fabriq-pdf-*' -and (Test-Path $Path)) {
            Remove-Item $Path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Profile Data Overlay (Phase 1)' {

    BeforeEach {
        Mock Show-Info    { }
        Mock Show-Warning { }
        Mock Show-Error   { }
        Mock Write-TelemetryEvent       { }
        Mock Write-KernelTelemetryEvent { }

        # GetFullPath: $env:TEMP may be an 8.3 short name (SZK-WI~1) while the
        # kernel normalizes through GetFullPath, which expands it.
        $script:pdf = [System.IO.Path]::GetFullPath((Join-Path $env:TEMP ("fabriq-pdf-{0}" -f ([guid]::NewGuid().ToString('N')))))
        $null = New-Item -ItemType Directory -Path $script:pdf -Force
        Clear-FabriqProfileDataContext
    }

    AfterEach {
        Clear-FabriqProfileDataContext
        Remove-OverlayTemp $script:pdf
    }

    Context 'Resolve-ModuleDataPath' {

        It 'has a real module CSV to test against' {
            Test-Path $script:ModuleCsv | Should -BeTrue
        }

        It 'returns the input unchanged and stays silent when no context is active' {
            $r = Resolve-ModuleDataPath -Path $script:ModuleCsv
            $r | Should -Be $script:ModuleCsv
            Should -Invoke Show-Info    -Exactly -Times 0
            Should -Invoke Show-Warning -Exactly -Times 0
        }

        It 'returns the profile copy (and says so once) when the file exists in the data folder' {
            $copy = Join-Path $script:pdf 'modules\taskbar_config\taskbar_list.csv'
            New-OverlayFile -Path $copy -Lines @('Enabled,TargetName', '1,a')
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf

            $r = Resolve-ModuleDataPath -Path $script:ModuleCsv
            $r | Should -Be $copy
            Should -Invoke Show-Info    -Exactly -Times 1
            Should -Invoke Show-Warning -Exactly -Times 0
        }

        It 'falls back to the module path with ONE warning when the file is missing from the data folder' {
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf

            $r1 = Resolve-ModuleDataPath -Path $script:ModuleCsv
            $r2 = Resolve-ModuleDataPath -Path $script:ModuleCsv
            $r1 | Should -Be $script:ModuleCsv
            $r2 | Should -Be $script:ModuleCsv
            Should -Invoke Show-Warning -Exactly -Times 1
            Should -Invoke Show-Info    -Exactly -Times 0
        }

        It 'repeats the display after the context is set again (new batch)' {
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf
            $null = Resolve-ModuleDataPath -Path $script:ModuleCsv
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf
            $null = Resolve-ModuleDataPath -Path $script:ModuleCsv
            Should -Invoke Show-Warning -Exactly -Times 2
        }

        It 'leaves paths outside <repo>\modules untouched under a context' {
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf
            $outside = Join-Path $env:TEMP 'not-a-module\something.csv'
            $r = Resolve-ModuleDataPath -Path $outside
            $r | Should -Be $outside
            Should -Invoke Show-Info    -Exactly -Times 0
            Should -Invoke Show-Warning -Exactly -Times 0
        }

        It 'normalizes ..\ cross-module references before resolving' {
            $copy = Join-Path $script:pdf 'modules\printer_driver_config\printer_list.csv'
            New-OverlayFile -Path $copy -Lines @('Enabled,TargetHost,PrinterName', '1,,P1')
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf

            $crossRef = Join-Path $script:RepoRoot 'modules\standard\printer_delete\..\printer_driver_config\printer_list.csv'
            $r = Resolve-ModuleDataPath -Path $crossRef
            $r | Should -Be $copy
        }

        It 'is idempotent for a path that already points into the data folder' {
            $copy = Join-Path $script:pdf 'modules\taskbar_config\taskbar_list.csv'
            New-OverlayFile -Path $copy -Lines @('Enabled,TargetName', '1,a')
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf

            $r = Resolve-ModuleDataPath -Path $copy
            $r | Should -Be $copy
            Should -Invoke Show-Info    -Exactly -Times 0
            Should -Invoke Show-Warning -Exactly -Times 0
        }
    }

    Context 'Import-ModuleCsv overlay hook' {

        It 'reads rows from the profile copy when present' {
            $copy = Join-Path $script:pdf 'modules\taskbar_config\taskbar_list.csv'
            New-OverlayFile -Path $copy -Lines @('Enabled,TargetName', '1,from-profile-a', '1,from-profile-b')
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf

            $r = Import-ModuleCsv -Path $script:ModuleCsv
            @($r).Count | Should -Be 2
            (@($r) | ForEach-Object TargetName) | Should -Be @('from-profile-a', 'from-profile-b')
        }

        It 'still applies the strict Segment filter to the profile copy' {
            $copy = Join-Path $script:pdf 'modules\taskbar_config\taskbar_list.csv'
            New-OverlayFile -Path $copy -Lines @('Enabled,TargetName,Segment', '1,a,alpha', '1,b,beta', '1,c,')
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf

            $r = Import-ModuleCsv -Path $script:ModuleCsv -Segment 'alpha'
            @($r).Count | Should -Be 1
            @($r)[0].TargetName | Should -Be 'a'
        }

        It 'reads the module copy when no context is active (backward compatibility)' {
            $r = Import-ModuleCsv -Path $script:ModuleCsv
            ($null -eq $r) | Should -BeFalse
            # The shipped taskbar_list.csv carries the AppId column; a
            # profile-copy fixture above never does.
            (@($r)[0].PSObject.Properties.Name -contains 'AppId') | Should -BeTrue
        }
    }

    Context 'Context lifecycle helpers' {

        It 'Get-FabriqProfileDataDir returns the sibling folder of a profile CSV when it exists' {
            $profileCsv = Join-Path $script:pdf 'Master_X.csv'
            New-OverlayFile -Path $profileCsv -Lines @('Order,ScriptPath,Enabled')
            $folder = Join-Path $script:pdf 'Master_X'
            $null = New-Item -ItemType Directory -Path $folder -Force

            Get-FabriqProfileDataDir -ProfilePath $profileCsv | Should -Be $folder
        }

        It 'Get-FabriqProfileDataDir returns an empty string when the folder does not exist' {
            $profileCsv = Join-Path $script:pdf 'Master_Y.csv'
            New-OverlayFile -Path $profileCsv -Lines @('Order,ScriptPath,Enabled')
            Get-FabriqProfileDataDir -ProfilePath $profileCsv | Should -Be ''
        }

        It 'Set-/Clear-FabriqProfileDataContext drive FABRIQ_PROFILE_DATA_DIR' {
            Set-FabriqProfileDataContext -ProfileDataDir $script:pdf
            $env:FABRIQ_PROFILE_DATA_DIR | Should -Be $script:pdf
            Clear-FabriqProfileDataContext
            [string]::IsNullOrEmpty($env:FABRIQ_PROFILE_DATA_DIR) | Should -BeTrue
        }
    }

    Context 'Test-FabriqResumeDataDir (fail-closed resume guard)' {

        It 'accepts a state written without a data folder' {
            $state = [pscustomobject]@{ ProfileDataDir = '' }
            Test-FabriqResumeDataDir -ResumeState $state | Should -BeTrue
        }

        It 'accepts a legacy state that has no ProfileDataDir field' {
            $state = [pscustomobject]@{ ProfilePath = 'C:\profiles\foo.csv' }
            Test-FabriqResumeDataDir -ResumeState $state | Should -BeTrue
        }

        It 'accepts a state whose data folder still exists' {
            $state = [pscustomobject]@{ ProfileDataDir = $script:pdf }
            Test-FabriqResumeDataDir -ResumeState $state | Should -BeTrue
            Should -Invoke Show-Error -Exactly -Times 0
        }

        It 'refuses (with an error) when the recorded data folder is gone' {
            $gone = Join-Path $script:pdf 'vanished'
            $state = [pscustomobject]@{ ProfileDataDir = $gone }
            Test-FabriqResumeDataDir -ResumeState $state | Should -BeFalse
            Should -Invoke Show-Error -Times 1
        }
    }
}
