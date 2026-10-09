# ========================================
# Pester v5 unit tests for windows_license_oem_install's decision logic
# ========================================
# Functions: modules/standard/windows_license_config/windows_license_oem_install.ps1 ::
#            Get-KeyPartial / Get-OemKeyPlan / Get-OemVerifyOutcome /
#            Get-MaskedKey
# Run      : powershell.exe -File ./dev/run_tests.ps1
#
# The module body runs on import, so the function definitions are lifted out
# with the AST and evaluated on their own (same approach as
# tests/modules/hostlist_precheck). Nothing outside the function bodies runs.
#
# Only the PURE functions are exercised. No test reads this machine's firmware
# or licensing state, so the result is the same on a developer PC and on CI.
# That matters here more than usual: a virtual machine has no firmware key, so
# the Success and Skipped paths of this script cannot be reached on the
# developer PC, on CI or on the test rig VM - these tests are the only
# automated cover they have.
#
# What is pinned is the direction every doubtful case falls in. The script
# promises that Success and Skipped always refer to the OEM key of this PC:
# a missing or malformed firmware key must stop the run, an unreadable
# license state must never count as "already installed", and after the
# install a different key must never pass as the firmware key. The last
# block pins what the script must not do at all: prompt for a key, read a
# CSV, or hand the raw key to anything that displays it.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    $script:ModuleDir = Join-Path $script:RepoRoot 'modules\standard\windows_license_config'
    $script:ModuleScript = Join-Path $script:ModuleDir 'windows_license_oem_install.ps1'
    $script:SiblingScript = Join-Path $script:ModuleDir 'windows_license_install.ps1'

    function Get-TestScriptAst {
        param([string]$Path)
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $Path, [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors -and $parseErrors.Count -gt 0) {
            throw "$(Split-Path $Path -Leaf) failed to parse: $($parseErrors[0].Message)"
        }
        return $ast
    }

    function Get-TestFunctionText {
        # Source text of a top-level function, with line endings normalized.
        param($Ast, [string]$Name)
        $fn = $Ast.FindAll(
            { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name },
            $false) | Select-Object -First 1
        if ($null -eq $fn) { return $null }
        return ($fn.Extent.Text -replace "`r`n", "`n")
    }

    $script:Ast = Get-TestScriptAst -Path $script:ModuleScript
    $script:SiblingAst = Get-TestScriptAst -Path $script:SiblingScript

    $functions = $script:Ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false)
    foreach ($fn in $functions) {
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    # Every command the script calls, by name (comments and strings excluded).
    $script:CommandNames = @($script:Ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.CommandAst] },
        $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })

    # Made-up keys: correct layout, not real product keys.
    $script:OemKey = 'AAAAA-BBBBB-CCCCC-DDDDD-EEEEE'
}

Describe 'Get-KeyPartial' {

    It 'returns the last 5 characters' {
        Get-KeyPartial 'AAAAA-BBBBB-CCCCC-DDDDD-EEEEE' | Should -BeExactly 'EEEEE'
    }

    It 'upper-cases the result, as Windows reports PartialProductKey' {
        Get-KeyPartial 'aaaaa-bbbbb-ccccc-ddddd-ab1cd' | Should -BeExactly 'AB1CD'
    }

    It 'ignores surrounding whitespace' {
        Get-KeyPartial '  AAAAA-BBBBB-CCCCC-DDDDD-EEEEE  ' | Should -BeExactly 'EEEEE'
    }

    It 'returns an empty string for <Label>' -ForEach @(
        @{ Label = 'a value shorter than 5 characters'; Key = 'ABCD' }
        @{ Label = 'an empty string'; Key = '' }
        @{ Label = 'null'; Key = $null }
    ) {
        Get-KeyPartial $Key | Should -BeExactly ''
    }
}

Describe 'Get-OemKeyPlan' {

    Context 'the firmware holds no usable key' {

        It 'returns NoKey for <Label>' -ForEach @(
            @{ Label = 'an empty string'; Key = '' }
            @{ Label = 'whitespace only'; Key = '   ' }
            @{ Label = 'null'; Key = $null }
        ) {
            Get-OemKeyPlan -OemKey $Key -InstalledPartials @('EEEEE') | Should -BeExactly 'NoKey'
        }

        It 'returns BadFormat for <Label>' -ForEach @(
            @{ Label = 'a key without dashes'; Key = 'AAAAABBBBBCCCCCDDDDDEEEEE' }
            @{ Label = 'a key with 4 groups'; Key = 'AAAAA-BBBBB-CCCCC-DDDDD' }
            @{ Label = 'a key with 6 groups'; Key = 'AAAAA-BBBBB-CCCCC-DDDDD-EEEEE-FFFFF' }
            @{ Label = 'a key with a symbol'; Key = 'AAAAA-BBBBB-CCCCC-DDDDD-EEEE!' }
            @{ Label = 'a key with a short group'; Key = 'AAAAA-BBBB-CCCCC-DDDDD-EEEEE' }
        ) {
            Get-OemKeyPlan -OemKey $Key -InstalledPartials @() | Should -BeExactly 'BadFormat'
        }

        It 'does not skip a malformed key just because its last 5 characters are installed' {
            Get-OemKeyPlan -OemKey 'EEEEE' -InstalledPartials @('EEEEE') | Should -BeExactly 'BadFormat'
        }
    }

    Context 'the OEM key is already installed' {

        It 'returns AlreadyInstalled when the installed key ends in the same 5 characters' {
            Get-OemKeyPlan -OemKey $script:OemKey -InstalledPartials @('EEEEE') | Should -BeExactly 'AlreadyInstalled'
        }

        It 'finds the OEM key among several installed keys' {
            Get-OemKeyPlan -OemKey $script:OemKey -InstalledPartials @('XXXXX', 'EEEEE') | Should -BeExactly 'AlreadyInstalled'
        }

        It 'compares without regard to case or surrounding whitespace' {
            Get-OemKeyPlan -OemKey ' aaaaa-bbbbb-ccccc-ddddd-eeeee ' -InstalledPartials @(' eeeee ') |
                Should -BeExactly 'AlreadyInstalled'
        }
    }

    Context 'the OEM key has to be installed' {

        It 'returns Install when a different key is installed' {
            Get-OemKeyPlan -OemKey $script:OemKey -InstalledPartials @('XXXXX') | Should -BeExactly 'Install'
        }

        It 'returns Install when <Label> (an unreadable state is never "already installed")' -ForEach @(
            @{ Label = 'no key is installed'; Partials = @() }
            @{ Label = 'the installed keys are null'; Partials = $null }
            @{ Label = 'the installed keys are blank'; Partials = @('', '  ') }
        ) {
            Get-OemKeyPlan -OemKey $script:OemKey -InstalledPartials $Partials | Should -BeExactly 'Install'
        }
    }
}

Describe 'Get-OemVerifyOutcome' {

    It 'returns Match when the installed key ends in the expected characters' {
        Get-OemVerifyOutcome -ExpectedPartial 'EEEEE' -ActualPartials @('EEEEE') | Should -BeExactly 'Match'
    }

    It 'returns Match when one of several installed keys is the expected one' {
        Get-OemVerifyOutcome -ExpectedPartial 'EEEEE' -ActualPartials @('XXXXX', 'EEEEE') | Should -BeExactly 'Match'
    }

    It 'compares without regard to case' {
        Get-OemVerifyOutcome -ExpectedPartial 'EEEEE' -ActualPartials @('eeeee') | Should -BeExactly 'Match'
    }

    It 'returns Mismatch when only another key is installed (the old key is still in effect)' {
        Get-OemVerifyOutcome -ExpectedPartial 'EEEEE' -ActualPartials @('XXXXX') | Should -BeExactly 'Mismatch'
    }

    It 'never matches an empty expectation' {
        Get-OemVerifyOutcome -ExpectedPartial '' -ActualPartials @('XXXXX') | Should -BeExactly 'Mismatch'
    }

    It 'returns Unreadable when <Label>' -ForEach @(
        @{ Label = 'no key can be read'; Partials = @() }
        @{ Label = 'the installed keys are null'; Partials = $null }
        @{ Label = 'the installed keys are blank'; Partials = @('', '  ') }
    ) {
        Get-OemVerifyOutcome -ExpectedPartial 'EEEEE' -ActualPartials $Partials | Should -BeExactly 'Unreadable'
    }
}

Describe 'Get-MaskedKey' {

    It 'keeps the dashes and the last 5 characters and masks the rest' {
        Get-MaskedKey 'AAAAA-BBBBB-CCCCC-DDDDD-EEEEE' | Should -BeExactly '*****-*****-*****-*****-EEEEE'
    }

    It 'masks a short value completely' {
        Get-MaskedKey 'ABCDE' | Should -BeExactly '*****'
    }

    It 'returns an empty string for an empty value' {
        Get-MaskedKey '' | Should -BeExactly ''
    }
}

Describe 'windows_license_oem_install.ps1 structure' {

    It 'never prompts for a key (no manual-input fallback)' {
        $script:CommandNames | Should -Not -Contain 'Read-Host'
    }

    It 'never reads a CSV (no fallback to license_key.csv)' {
        $script:CommandNames | Should -Not -Contain 'Import-ModuleCsv'
        $script:CommandNames | Should -Not -Contain 'Import-Csv'
        $script:CommandNames | Should -Not -Contain 'Resolve-ModuleDataPath'
    }

    It 'asks for confirmation before installing' {
        $script:CommandNames | Should -Contain 'Confirm-ModuleExecution'
    }

    It 'hands the raw firmware key only to the helper functions and to InstallProductKey' {
        # Every use of $oemKey in the script body must sit directly inside one
        # of these commands, or be the assignment that fills it. Anything else
        # (Write-Host, Show-*, New-ModuleResult, a file write) would put the
        # full key on screen, in the transcript or in the execution history.
        $allowed = @('Get-MaskedKey', 'Get-KeyPartial', 'Get-OemKeyPlan', 'Invoke-CimMethod')

        $uses = @($script:Ast.FindAll({ param($node)
            if ($node -isnot [System.Management.Automation.Language.VariableExpressionAst]) { return $false }
            if ($node.VariablePath.UserPath -ne 'oemKey') { return $false }
            # Skip the helper functions: their own parameters are not the script's variable.
            $up = $node.Parent
            while ($null -ne $up) {
                if ($up -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $false }
                $up = $up.Parent
            }
            return $true
        }, $true))

        $uses.Count | Should -BeGreaterThan 1

        $violations = @()
        foreach ($use in $uses) {
            $owner = $use.Parent
            while ($null -ne $owner -and
                   $owner -isnot [System.Management.Automation.Language.CommandAst] -and
                   $owner -isnot [System.Management.Automation.Language.AssignmentStatementAst]) {
                $owner = $owner.Parent
            }

            $ok = $false
            if ($owner -is [System.Management.Automation.Language.CommandAst]) {
                $ok = $allowed -contains $owner.GetCommandName()
            }
            elseif ($owner -is [System.Management.Automation.Language.AssignmentStatementAst]) {
                $ok = [object]::ReferenceEquals($owner.Left, $use)
            }

            if (-not $ok) {
                $violations += "line $($use.Extent.StartLineNumber): $($use.Parent.Extent.Text)"
            }
        }

        $violations | Should -BeNullOrEmpty
    }

    It 'keeps <Name> identical to windows_license_install.ps1' -ForEach @(
        @{ Name = 'Get-MaskedKey' }
        @{ Name = 'Get-LicenseStatusText' }
    ) {
        $mine = Get-TestFunctionText -Ast $script:Ast -Name $Name
        $sibling = Get-TestFunctionText -Ast $script:SiblingAst -Name $Name

        $mine | Should -Not -BeNullOrEmpty
        $mine | Should -BeExactly $sibling
    }
}
