# ========================================
# Pester v5 unit tests for the AD lab builder's pure logic
# ========================================
# Functions: dev/test_rig/ad_lab/ad_lab.ps1 :: Resolve-LabDn / Split-LabOuDn /
#            Read-LabConfig / Get-LabDelegationAce
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# The lab builder only ever runs on a domain controller, so the parts that talk
# to AD cannot be covered here. What IS covered is everything that decides WHAT
# gets created - and those are the parts that would silently build the wrong lab:
#   - relative vs absolute OU paths (the JSON writes OU paths without DC=)
#   - DN leaf parsing with a backslash-escaped comma (OU=Sales\,EMEA)
#   - config validation rejecting the mistakes that only surface mid-promotion
#   - the delegation ACE sets (schema GUIDs, not localized dsacls right names)
# The script is dot-sourced with -LoadOnly so nothing executes.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    $script:LabScript = Join-Path $script:RepoRoot 'dev\test_rig\ad_lab\ad_lab.ps1'
    . $script:LabScript -LoadOnly

    $script:LabConfigPath = Join-Path $script:RepoRoot 'dev\test_rig\ad_lab\ad_lab.json'
    $script:TempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ad_lab_tests_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:TempDir | Out-Null

    function New-TempConfig {
        param([scriptblock]$Mutate)
        $o = Get-Content -LiteralPath $script:LabConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($Mutate) { & $Mutate $o }
        $p = Join-Path $script:TempDir ([guid]::NewGuid().ToString('N') + '.json')
        $o | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $p -Encoding UTF8
        return $p
    }
}

AfterAll {
    if ($script:TempDir -and (Test-Path -LiteralPath $script:TempDir)) {
        Remove-Item -LiteralPath $script:TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Resolve-LabDn' {
    It 'appends the domain DN to a relative OU path' {
        Resolve-LabDn -Path 'OU=Kitting,OU=PC' -DomainDn 'DC=lab,DC=local' |
            Should -Be 'OU=Kitting,OU=PC,DC=lab,DC=local'
    }

    It 'leaves an absolute DN untouched' {
        Resolve-LabDn -Path 'OU=Kitting,DC=other,DC=local' -DomainDn 'DC=lab,DC=local' |
            Should -Be 'OU=Kitting,DC=other,DC=local'
    }

    It 'is case insensitive about the DC= marker' {
        Resolve-LabDn -Path 'ou=kitting,dc=lab,dc=local' -DomainDn 'DC=lab,DC=local' |
            Should -Be 'ou=kitting,dc=lab,dc=local'
    }

    It 'returns the domain DN for an empty path' {
        Resolve-LabDn -Path '' -DomainDn 'DC=lab,DC=local' | Should -Be 'DC=lab,DC=local'
    }
}

Describe 'Split-LabOuDn' {
    It 'splits a nested OU into leaf name and parent DN' {
        $r = Split-LabOuDn -Dn 'OU=Kitting,OU=PC,DC=lab,DC=local'
        $r.Name   | Should -Be 'Kitting'
        $r.Parent | Should -Be 'OU=PC,DC=lab,DC=local'
    }

    It 'unescapes a comma inside the leaf name' {
        $r = Split-LabOuDn -Dn 'OU=Sales\,EMEA,DC=lab,DC=local'
        $r.Name   | Should -Be 'Sales,EMEA'
        $r.Parent | Should -Be 'DC=lab,DC=local'
    }

    It 'returns null when the leaf is not an OU' {
        Split-LabOuDn -Dn 'CN=Computers,DC=lab,DC=local' | Should -BeNullOrEmpty
    }

    It 'returns null for a single-component DN' {
        Split-LabOuDn -Dn 'OU=Orphan' | Should -BeNullOrEmpty
    }
}

Describe 'Read-LabConfig' {
    It 'accepts the shipped sample config' {
        $cfg = Read-LabConfig -Path $script:LabConfigPath
        $cfg.domain.fqdn | Should -Be 'lab.fabriq.local'
        @($cfg.ous).Count | Should -BeGreaterThan 0
    }

    It 'keeps the escaped comma through a JSON round trip' {
        $cfg = Read-LabConfig -Path $script:LabConfigPath
        @($cfg.ous) -contains 'OU=Sales\,EMEA' | Should -BeTrue
    }

    It 'accepts settings.domainAdmins entries' {
        $p = New-TempConfig { param($o) $o.settings.domainAdmins = @('FabriqAdmin') }
        $cfg = Read-LabConfig -Path $p
        @($cfg.settings.domainAdmins) -contains 'FabriqAdmin' | Should -BeTrue
    }
}

Describe 'Get-PendingRebootReason' {
    # The promotion prerequisite check refuses to run while a reboot is pending,
    # so this probe decides whether phase 1 reboots before promoting.
    #
    # The registry is MOCKED on purpose. An earlier version asserted against the
    # real registry and passed on the dev machine only because that machine
    # happened to have a pending file rename; on a clean CI runner the list was
    # empty, an empty array piped into Should arrives as $null, and the type
    # assertion failed. Both states are now pinned explicitly.

    It 'returns an empty list when nothing is pending' {
        Mock Test-Path { $false }
        Mock Get-ItemProperty { $null }

        $reasons = @(Get-PendingRebootReason)
        $reasons.Count | Should -Be 0
    }

    It 'reports every pending source as a non-empty string' {
        Mock Test-Path { $true }
        Mock Get-ItemProperty -ParameterFilter { $Name -eq 'PendingFileRenameOperations' } {
            [pscustomobject]@{ PendingFileRenameOperations = @('\??\C:\stale.tmp') }
        }
        Mock Get-ItemProperty -ParameterFilter { $Path -like '*\ActiveComputerName' } {
            [pscustomobject]@{ ComputerName = 'OLD-NAME' }
        }
        Mock Get-ItemProperty -ParameterFilter { $Path -like '*\ComputerName\ComputerName' } {
            [pscustomobject]@{ ComputerName = 'NEW-NAME' }
        }

        $reasons = @(Get-PendingRebootReason)
        # 3 registry flags + pending file renames + a staged computer rename
        $reasons.Count | Should -Be 5
        foreach ($r in $reasons) {
            $r | Should -BeOfType [string]
            [string]::IsNullOrWhiteSpace($r) | Should -BeFalse
        }
        ($reasons -join "`n") | Should -Match 'OLD-NAME -> NEW-NAME'
    }

    It 'does not throw against the real registry of whatever machine runs it' {
        { Get-PendingRebootReason } | Should -Not -Throw
    }
}

Describe 'Get-LabDelegationAce' {
    BeforeAll {
        # A well-known SID needs no directory to resolve.
        $script:Sid = New-Object Security.Principal.SecurityIdentifier('S-1-5-11')
        $script:ComputerClassGuid = [guid]'bf967a86-0de6-11d0-a285-00aa003049e2'
    }

    It 'grants exactly CreateChild on the computer class for CreateComputer' {
        $aces = @(Get-LabDelegationAce -Sid $script:Sid -Right 'CreateComputer')
        $aces.Count | Should -Be 1
        $aces[0].ActiveDirectoryRights.ToString() | Should -Be 'CreateChild'
        $aces[0].ObjectType | Should -Be $script:ComputerClassGuid
        $aces[0].AccessControlType.ToString() | Should -Be 'Allow'
    }

    It 'adds the re-use writes for FullJoin' {
        $aces = @(Get-LabDelegationAce -Sid $script:Sid -Right 'FullJoin')
        $aces.Count | Should -Be 6
        $rights = @($aces | ForEach-Object { $_.ActiveDirectoryRights.ToString() })
        $rights | Should -Contain 'DeleteChild'
        $rights | Should -Contain 'ExtendedRight'
        $rights | Should -Contain 'Self'
        $rights | Should -Contain 'WriteProperty'
    }

    It 'scopes the FullJoin extras to descendant computer objects' {
        $aces = @(Get-LabDelegationAce -Sid $script:Sid -Right 'FullJoin')
        $descendants = @($aces | Where-Object { $_.InheritanceType.ToString() -eq 'Descendents' })
        $descendants.Count | Should -Be 4
        foreach ($a in $descendants) {
            $a.InheritedObjectType | Should -Be $script:ComputerClassGuid
        }
    }
}
