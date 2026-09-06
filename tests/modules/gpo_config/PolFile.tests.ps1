# ========================================
# Pester v5 unit tests for modules/standard/gpo_config/lib/PolFile.ps1
# ========================================
# Functions: Read-PolFile / Write-PolFile / ConvertTo-PolData /
#            ConvertFrom-PolData / ConvertTo-PolEntryFromAction /
#            Merge-PolEntries / Update-GptIniVersion / identity helpers
# Run    : powershell.exe -File ./dev/run_tests.ps1
#
# Pins the MS-GPREG "PReg" codec that gpo_config / gpo_backup rely on.
# The fixture sample_machine.pol is a real Machine\Registry.pol produced by
# gpedit.msc ("Configure Automatic Updates" = Disabled + "No auto-restart
# with logged on users" = Enabled): 12 entries, 10 of them **del. markers.
# A regression here silently produces a Registry.pol that gpupdate ignores
# or that deletes the wrong values - not observable in static analysis.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'modules\standard\gpo_config\lib\PolFile.ps1')
    $script:Fixture = Join-Path $PSScriptRoot 'fixtures\sample_machine.pol'
}

Describe 'Read-PolFile' {

    It 'returns an empty list for a missing file' {
        $r = Read-PolFile -Path (Join-Path $TestDrive 'nope.pol')
        @($r).Count | Should -Be 0
    }

    It 'returns an empty list for a zero-byte file' {
        $p = Join-Path $TestDrive 'empty.pol'
        [IO.File]::WriteAllBytes($p, @())
        @(Read-PolFile -Path $p).Count | Should -Be 0
    }

    It 'parses the real gpedit sample (12 entries, 10 **del. markers)' {
        $entries = @(Read-PolFile -Path $script:Fixture)
        $entries.Count | Should -Be 12
        $entries[0].Key       | Should -Be 'Software\Policies\Microsoft\Windows\WindowsUpdate\AU'
        $entries[0].ValueName | Should -Be 'NoAutoUpdate'
        $entries[0].Type      | Should -Be 4
        ConvertFrom-PolData -Type $entries[0].Type -Data $entries[0].Data | Should -Be '1'
        @($entries | Where-Object { $_.ValueName -like '**del.*' }).Count | Should -Be 10
        $entries[11].ValueName | Should -Be 'NoAutoRebootWithLoggedOnUsers'
    }

    It 'throws on a file without the PReg signature' {
        $p = Join-Path $TestDrive 'bad.pol'
        [IO.File]::WriteAllBytes($p, [Text.Encoding]::ASCII.GetBytes('XXXX') + @(1,0,0,0))
        { Read-PolFile -Path $p } | Should -Throw '*PReg*'
    }

    It 'throws on a truncated record instead of returning partial data' {
        $all = [IO.File]::ReadAllBytes($script:Fixture)
        $p = Join-Path $TestDrive 'trunc.pol'
        [IO.File]::WriteAllBytes($p, $all[0..($all.Length - 20)])
        { Read-PolFile -Path $p } | Should -Throw
    }
}

Describe 'Write-PolFile round-trip' {

    It 'reproduces the gpedit sample byte-for-byte' {
        $entries = @(Read-PolFile -Path $script:Fixture)
        $out = Join-Path $TestDrive 'rt.pol'
        Write-PolFile -Path $out -Entries $entries
        (Get-FileHash $out -Algorithm SHA256).Hash | Should -Be (Get-FileHash $script:Fixture -Algorithm SHA256).Hash
    }

    It 'writes the PReg header and version 1' {
        $out = Join-Path $TestDrive 'hdr.pol'
        Write-PolFile -Path $out -Entries @()
        $b = [IO.File]::ReadAllBytes($out)
        $b.Length | Should -Be 8
        [Text.Encoding]::ASCII.GetString($b, 0, 4) | Should -Be 'PReg'
        [BitConverter]::ToInt32($b, 4) | Should -Be 1
    }

    It 'creates the parent directory and leaves no temp file behind' {
        $out = Join-Path $TestDrive 'sub\dir\Registry.pol'
        Write-PolFile -Path $out -Entries @((New-PolEntry -Key 'Software\X' -ValueName 'V' -Type 4 -Data (ConvertTo-PolData -Type 4 -Value '1')))
        Test-Path $out | Should -BeTrue
        Test-Path "$out.fabriq.tmp" | Should -BeFalse
        @(Read-PolFile -Path $out).Count | Should -Be 1
    }

    It 'round-trips every supported type' {
        $src = @(
            (New-PolEntry -Key 'Software\T' -ValueName 'dw'  -Type 4  -Data (ConvertTo-PolData -Type 4  -Value '0xFFFFFFFF')),
            (New-PolEntry -Key 'Software\T' -ValueName 'qw'  -Type 11 -Data (ConvertTo-PolData -Type 11 -Value '1234567890123')),
            (New-PolEntry -Key 'Software\T' -ValueName 'sz'  -Type 1  -Data (ConvertTo-PolData -Type 1  -Value 'hello world')),
            (New-PolEntry -Key 'Software\T' -ValueName 'ex'  -Type 2  -Data (ConvertTo-PolData -Type 2  -Value '%SystemRoot%\x')),
            (New-PolEntry -Key 'Software\T' -ValueName 'ms'  -Type 7  -Data (ConvertTo-PolData -Type 7  -Value 'a|b|c')),
            (New-PolEntry -Key 'Software\T' -ValueName 'bin' -Type 3  -Data (ConvertTo-PolData -Type 3  -Value '00 FF 10'))
        )
        $out = Join-Path $TestDrive 'types.pol'
        Write-PolFile -Path $out -Entries $src
        $back = @(Read-PolFile -Path $out)
        $back.Count | Should -Be 6
        for ($i = 0; $i -lt 6; $i++) { Test-PolEntryEqual -A $src[$i] -B $back[$i] | Should -BeTrue }
        ConvertFrom-PolData -Type 4  -Data $back[0].Data | Should -Be '4294967295'
        ConvertFrom-PolData -Type 11 -Data $back[1].Data | Should -Be '1234567890123'
        ConvertFrom-PolData -Type 1  -Data $back[2].Data | Should -Be 'hello world'
        ConvertFrom-PolData -Type 2  -Data $back[3].Data | Should -Be '%SystemRoot%\x'
        ConvertFrom-PolData -Type 7  -Data $back[4].Data | Should -Be 'a|b|c'
        ConvertFrom-PolData -Type 3  -Data $back[5].Data | Should -Be '00FF10'
    }
}

Describe 'ConvertTo-PolData' {

    It 'encodes REG_DWORD as 4 little-endian bytes (decimal, hex, negative)' {
        (ConvertTo-PolData -Type 4 -Value '1')          | Should -Be @(1,0,0,0)
        (ConvertTo-PolData -Type 4 -Value '0x10')       | Should -Be @(16,0,0,0)
        (ConvertTo-PolData -Type 4 -Value '4294967295') | Should -Be @(255,255,255,255)
        (ConvertTo-PolData -Type 4 -Value '-1')         | Should -Be @(255,255,255,255)
    }

    It 'rejects non-numeric REG_DWORD input' {
        { ConvertTo-PolData -Type 4 -Value 'abc' } | Should -Throw
    }

    It 'NUL-terminates strings and double-terminates MULTI_SZ' {
        (ConvertTo-PolData -Type 1 -Value 'ab') | Should -Be @(97,0,98,0,0,0)
        (ConvertTo-PolData -Type 7 -Value 'a|b') | Should -Be @(97,0,0,0,98,0,0,0,0,0)
        (ConvertTo-PolData -Type 7 -Value '')    | Should -Be @(0,0)
    }

    It 'rejects odd-length REG_BINARY hex' {
        { ConvertTo-PolData -Type 3 -Value 'ABC' } | Should -Throw
    }
}

Describe 'ConvertTo-PolEntryFromAction' {

    It 'Delete produces a **del. marker with a single-space REG_SZ payload' {
        $e = ConvertTo-PolEntryFromAction -Key 'Software\K' -ValueName 'X' -Action 'Delete'
        $e.ValueName | Should -Be '**del.X'
        $e.Type      | Should -Be 1
        $e.Data      | Should -Be @(32,0,0,0)
    }

    It 'DeleteAllValues produces **delvals.' {
        (ConvertTo-PolEntryFromAction -Key 'Software\K' -Action 'DeleteAllValues').ValueName | Should -Be '**delvals.'
    }

    It 'CreateKey produces an empty-name REG_NONE entry' {
        $e = ConvertTo-PolEntryFromAction -Key 'Software\K' -Action 'CreateKey'
        $e.ValueName | Should -Be ''
        $e.Type      | Should -Be 0
        $e.Data.Length | Should -Be 0
    }

    It 'Unmanage produces $null' {
        ConvertTo-PolEntryFromAction -Key 'Software\K' -ValueName 'X' -Action 'Unmanage' | Should -BeNullOrEmpty
    }

    It 'Set rejects unknown types' {
        { ConvertTo-PolEntryFromAction -Key 'Software\K' -ValueName 'X' -Action 'Set' -Type 'REG_FOO' -Value '1' } | Should -Throw
    }

    It 'Get-PolEntryAction inverts the mapping' {
        Get-PolEntryAction -Entry (ConvertTo-PolEntryFromAction -Key 'K' -ValueName 'X' -Action 'Delete')       | Should -Be 'Delete'
        Get-PolEntryAction -Entry (ConvertTo-PolEntryFromAction -Key 'K' -Action 'DeleteAllValues')             | Should -Be 'DeleteAllValues'
        Get-PolEntryAction -Entry (ConvertTo-PolEntryFromAction -Key 'K' -Action 'CreateKey')                   | Should -Be 'CreateKey'
        Get-PolEntryAction -Entry (New-PolEntry -Key 'K' -ValueName 'V' -Type 4 -Data @(1,0,0,0))               | Should -Be 'Set'
        Get-PolEntryAction -Entry (New-PolEntry -Key 'K' -ValueName '**DeleteKeys' -Type 1 -Data @(32,0,0,0))   | Should -Be 'Unsupported'
    }
}

Describe 'Identity and merge' {

    It 'treats X and **del.X as the same identity (case-insensitive)' {
        (Get-PolEntryIdentity -Key 'Software\K' -ValueName 'X') | Should -Be (Get-PolEntryIdentity -Key 'software\k' -ValueName '**del.x')
        (Get-PolEntryIdentity -Key 'Software\K' -ValueName 'X') | Should -Not -Be (Get-PolEntryIdentity -Key 'Software\K' -ValueName 'Y')
    }

    It 'Set replaces an existing **del. marker and preserves untouched entries in order' {
        $existing = @(
            (New-PolEntry -Key 'Software\K' -ValueName 'A' -Type 4 -Data @(1,0,0,0)),
            (ConvertTo-PolEntryFromAction -Key 'Software\K' -ValueName 'B' -Action 'Delete'),
            (New-PolEntry -Key 'Software\K' -ValueName 'C' -Type 4 -Data @(3,0,0,0))
        )
        $apply = @((New-PolEntry -Key 'Software\K' -ValueName 'B' -Type 4 -Data @(2,0,0,0)))
        $m = @(Merge-PolEntries -Existing $existing -Apply $apply)
        $m.Count | Should -Be 3
        $m[0].ValueName | Should -Be 'A'
        $m[1].ValueName | Should -Be 'C'
        $m[2].ValueName | Should -Be 'B'
        $m[2].Type      | Should -Be 4
    }

    It 'RemoveIdentities drops entries (Unmanage) without adding anything' {
        $existing = @(
            (New-PolEntry -Key 'Software\K' -ValueName 'A' -Type 4 -Data @(1,0,0,0)),
            (New-PolEntry -Key 'Software\K' -ValueName 'B' -Type 4 -Data @(2,0,0,0))
        )
        $m = @(Merge-PolEntries -Existing $existing -RemoveIdentities @((Get-PolEntryIdentity -Key 'Software\K' -ValueName 'A')) -Apply @())
        $m.Count | Should -Be 1
        $m[0].ValueName | Should -Be 'B'
    }

    It 'merging the gpedit sample onto itself is a no-op in content' {
        $existing = @(Read-PolFile -Path $script:Fixture)
        $m = @(Merge-PolEntries -Existing $existing -Apply $existing)
        $m.Count | Should -Be 12
        for ($i = 0; $i -lt 12; $i++) { Test-PolEntryEqual -A $existing[$i] -B $m[$i] | Should -BeTrue }
    }

    It 'returns an empty array (not $null-wrapped) when everything is removed' {
        $existing = @((New-PolEntry -Key 'Software\K' -ValueName 'A' -Type 4 -Data @(1,0,0,0)))
        $m = @(Merge-PolEntries -Existing $existing -RemoveIdentities @((Get-PolEntryIdentity -Key 'Software\K' -ValueName 'A')) -Apply @())
        $m.Count | Should -Be 0
    }
}

Describe 'Update-GptIniVersion' {

    It 'creates gpt.ini with the Registry CSE + Admin Templates GUIDs when missing (Machine -> low 16 bits)' {
        $p = Join-Path $TestDrive 'new\gpt.ini'
        $r = Update-GptIniVersion -Path $p -Scope 'Machine'
        $r.Old | Should -Be 0
        $r.New | Should -Be 1
        $txt = Get-Content $p
        $txt[0] | Should -Be '[General]'
        ($txt -join "`n") | Should -Match 'gPCMachineExtensionNames=\[\{35378EAC-683F-11D2-A89A-00C04FBBCFA2\}\{D02B1F72-3407-48AE-BA88-E8213C6761F1\}\]'
        ($txt -join "`n") | Should -Match 'Version=1'
    }

    It 'User bumps the high 16 bits and writes the user GUID line' {
        $p = Join-Path $TestDrive 'user\gpt.ini'
        $r = Update-GptIniVersion -Path $p -Scope 'User'
        $r.New | Should -Be 65536
        (Get-Content $p -Raw) | Should -Match 'gPCUserExtensionNames=\[\{35378EAC-683F-11D2-A89A-00C04FBBCFA2\}\{D02B1F73-3407-48AE-BA88-E8213C6761F1\}\]'
    }

    It 'increments only the targeted half and preserves the other extension line' {
        $p = Join-Path $TestDrive 'both\gpt.ini'
        New-Item -ItemType Directory -Path (Split-Path $p) -Force | Out-Null
        [IO.File]::WriteAllLines($p, [string[]]@(
            '[General]',
            'gPCUserExtensionNames=[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F73-3407-48AE-BA88-E8213C6761F1}]',
            'Version=131074',
            'gPCMachineExtensionNames=[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F72-3407-48AE-BA88-E8213C6761F1}]'
        ))
        (Update-GptIniVersion -Path $p -Scope 'Machine').New | Should -Be 131075
        (Update-GptIniVersion -Path $p -Scope 'User').New    | Should -Be 196611
        $raw = Get-Content $p -Raw
        $raw | Should -Match 'gPCUserExtensionNames=\[\{35378EAC-683F-11D2-A89A-00C04FBBCFA2\}\{D02B1F73-3407-48AE-BA88-E8213C6761F1\}\]'
        $raw | Should -Match 'gPCMachineExtensionNames=\[\{35378EAC-683F-11D2-A89A-00C04FBBCFA2\}\{D02B1F72-3407-48AE-BA88-E8213C6761F1\}\]'
        @(Get-Content $p | Where-Object { $_ -match '^Version=' }).Count | Should -Be 1
    }

    It 'adds the Admin Templates GUID to an existing CSE group and keeps foreign groups' {
        $p = Join-Path $TestDrive 'guid\gpt.ini'
        New-Item -ItemType Directory -Path (Split-Path $p) -Force | Out-Null
        [IO.File]::WriteAllLines($p, [string[]]@(
            '[General]',
            'gPCMachineExtensionNames=[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}][{827D319E-6EAC-11D2-A4EA-00C04F79F83A}{803E14A0-B4FB-11D0-A0D0-00A0C90F574B}]',
            'Version=5'
        ))
        $null = Update-GptIniVersion -Path $p -Scope 'Machine'
        $line = Get-Content $p | Where-Object { $_ -match '^gPCMachineExtensionNames=' }
        $line | Should -Be 'gPCMachineExtensionNames=[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F72-3407-48AE-BA88-E8213C6761F1}][{827D319E-6EAC-11D2-A4EA-00C04F79F83A}{803E14A0-B4FB-11D0-A0D0-00A0C90F574B}]'
    }

    It 'wraps the 16-bit counter instead of overflowing into the other half' {
        $p = Join-Path $TestDrive 'wrap\gpt.ini'
        New-Item -ItemType Directory -Path (Split-Path $p) -Force | Out-Null
        [IO.File]::WriteAllLines($p, [string[]]@('[General]', 'Version=65535'))
        (Update-GptIniVersion -Path $p -Scope 'Machine').New | Should -Be 0
    }
}
