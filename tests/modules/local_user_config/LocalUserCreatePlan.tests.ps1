# ========================================
# Pester v5 unit tests for local_user_config's per-row creation plan
# ========================================
# Function: modules/standard/local_user_config/local_user_config.ps1 ::
#           Get-LocalUserCreatePlan
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# The module body runs on import (a shipped module has no -LoadOnly switch
# and adding one would change its contract), so the function definitions are
# lifted out with the AST and evaluated on their own. Nothing outside the
# function bodies is executed.
#
# What is covered is the decision table that turns one CSV row into "create
# with this password", "create WITHOUT a password" or "refuse". An empty
# Password cell is a deliberate request for an account without a password,
# so the two inputs that only look like one must stay apart: a line that
# ends before the Password column (a truncated row) and a row without a
# user name (an unfilled local_user_setup placeholder). Reading either as
# "empty" would hand out an account nobody asked for.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    $script:ModuleScript = Join-Path $script:RepoRoot 'modules\standard\local_user_config\local_user_config.ps1'

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:ModuleScript, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "local_user_config.ps1 failed to parse: $($parseErrors[0].Message)"
    }

    $script:Functions = $ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false)
    foreach ($fn in $script:Functions) {
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    $script:ModuleText = Get-Content -LiteralPath $script:ModuleScript -Raw

    $script:CsvHeader = 'Enabled,UserName,Password,PasswordNeverExpires,UserMayNotChangePassword,Group,Description,Segment'

    # Rows go through the real CSV parser, so the tests see exactly what
    # Import-Csv hands the module (empty cell = '', missing field = $null).
    function ConvertTo-UserRow {
        param([string]$Line)
        return @("$script:CsvHeader`r`n$Line" | ConvertFrom-Csv)[0]
    }
}

Describe 'Get-LocalUserCreatePlan' {

    Context 'Password cell' {

        It 'plans a normal creation when the Password cell has a value' {
            $r = Get-LocalUserCreatePlan -Row (ConvertTo-UserRow '1,alice,Pw1234,1,1,Users,,')
            $r.Error | Should -BeNullOrEmpty
            $r.Mode  | Should -Be 'Password'
        }

        It 'plans a creation without a password when the Password cell is empty' {
            $r = Get-LocalUserCreatePlan -Row (ConvertTo-UserRow '1,kiosk,,1,1,Users,,')
            $r.Error | Should -BeNullOrEmpty
            $r.Mode  | Should -Be 'NoPassword'
        }

        It 'reads an unquoted whitespace-only cell as empty (the CSV parser drops it)' {
            # Pins the parser behaviour the "empty = no password" rule sits
            # on: such a cell never reaches the module as a space.
            $row = ConvertTo-UserRow '1,kiosk, ,1,1,Users,,'
            $row.Password | Should -Be ''
            (Get-LocalUserCreatePlan -Row $row).Mode | Should -Be 'NoPassword'
        }

        It 'keeps a quoted whitespace password as a password' {
            $row = ConvertTo-UserRow '1,alice," ",1,1,Users,,'
            $row.Password | Should -Be ' '
            (Get-LocalUserCreatePlan -Row $row).Mode | Should -Be 'Password'
        }

        It 'does not judge the value: an undecrypted ENC: string is still a password' {
            $r = Get-LocalUserCreatePlan -Row (ConvertTo-UserRow '1,alice,ENC:Zm9vYmFy,1,1,Users,,')
            $r.Error | Should -BeNullOrEmpty
            $r.Mode  | Should -Be 'Password'
        }
    }

    Context 'Rows that must be refused' {

        It 'refuses a line that ends before the Password column' {
            # "1,bob" - Import-Csv yields $null (not '') for the missing
            # fields. That is a broken line, not a request for no password.
            $row = ConvertTo-UserRow '1,bob'
            $row.Password | Should -BeNullOrEmpty
            $null -eq $row.Password | Should -BeTrue

            $r = Get-LocalUserCreatePlan -Row $row
            $r.Mode  | Should -BeNullOrEmpty
            $r.Error | Should -Match 'Password field is missing'
        }

        It 'refuses a row whose object has no Password property at all' {
            $row = [PSCustomObject]@{ Enabled = '1'; UserName = 'bob'; Group = 'Users' }
            (Get-LocalUserCreatePlan -Row $row).Error | Should -Match 'Password field is missing'
        }

        It 'refuses an unfilled placeholder row (no user name, no password)' {
            # apps/local_user_setup leaves such rows until the operator
            # registers a name. They must never reach New-LocalUser.
            $r = Get-LocalUserCreatePlan -Row (ConvertTo-UserRow '1,,,1,1,Users,,')
            $r.Mode  | Should -BeNullOrEmpty
            $r.Error | Should -Match 'UserName is empty'
        }

        It 'refuses a whitespace-only user name' {
            $r = Get-LocalUserCreatePlan -Row (ConvertTo-UserRow '1,"  ",Pw1234,1,1,Users,,')
            $r.Error | Should -Match 'UserName is empty'
        }

        It 'reports the missing user name before the missing password' {
            $row = [PSCustomObject]@{ Enabled = '1'; UserName = ''; Group = 'Users' }
            (Get-LocalUserCreatePlan -Row $row).Error | Should -Match 'UserName is empty'
        }

        It 'refuses a null row' {
            (Get-LocalUserCreatePlan -Row $null).Error | Should -Match 'UserName is empty'
        }
    }

    Context 'Per-PC rows (local_user_host_list.csv shape)' {

        It 'applies the same rule to a per-PC row' {
            $hostRow = @(
                "Enabled,NewPCName,UserName,Password,PasswordNeverExpires,UserMayNotChangePassword,Group,Description,AutoLogon,Segment`r`n" +
                '1,NEW-PC-01,pc01_kiosk,,1,1,Users,,1,' | ConvertFrom-Csv)[0]
            $r = Get-LocalUserCreatePlan -Row $hostRow
            $r.Error | Should -BeNullOrEmpty
            $r.Mode  | Should -Be 'NoPassword'
        }
    }
}

Describe 'local_user_config.ps1 creation path' {

    It 'never passes -NoPassword to New-LocalUser' {
        # Measured 2026-10-05: -NoPassword leaves "user must change password
        # at next logon" set and has no -PasswordNeverExpires. The blank
        # password has to go in as an empty SecureString instead.
        $script:ModuleText | Should -Not -Match 'New-LocalUser[^\r\n]*-NoPassword'
        $script:ModuleText | Should -Not -Match 'NoPassword\s*=\s*\$true'
    }

    It 'builds the blank password as an empty SecureString' {
        $script:ModuleText | Should -Match 'New-Object\s+System\.Security\.SecureString'
    }

    It 'classifies the row before the existence check in the creation loop' {
        # A broken row must fail the same way whatever accounts the machine
        # already has, so the plan is consulted before Get-LocalUser.
        $loopStart = $script:ModuleText.IndexOf('# User Creation Process')
        $loopStart | Should -BeGreaterThan 0
        $loopText  = $script:ModuleText.Substring($loopStart)

        $planAt   = $loopText.IndexOf('Get-LocalUserCreatePlan')
        $existsAt = $loopText.IndexOf('Get-LocalUser -Name')
        $planAt   | Should -BeGreaterThan 0
        $existsAt | Should -BeGreaterThan $planAt
    }
}
