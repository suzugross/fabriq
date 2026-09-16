# ========================================
# Pester v5 unit tests for autologon_config's host-bound target selection
# ========================================
# Function: modules/standard/autologon_config/autologon_config.ps1 ::
#           Get-HostAutoLogonTarget
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# The module body runs on import (a shipped module has no -LoadOnly switch
# and adding one would change its contract), so the function definitions are
# lifted out with the AST and evaluated on their own. Nothing outside the
# function bodies is executed.
#
# What is covered is the decision table behind __AUTO_to_[HOSTUSER]__: which
# per-PC local user the next boot logs on as. Getting it wrong arms an
# autologon for the wrong (or a non-existent) account, and the machine then
# parks on the logon screen where an unattended profile resume cannot go on.
# Every ambiguous input must therefore come back as an Error rather than a
# guess - the operator on site cannot be asked which account was meant.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    $script:ModuleScript = Join-Path $script:RepoRoot 'modules\standard\autologon_config\autologon_config.ps1'

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:ModuleScript, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "autologon_config.ps1 failed to parse: $($parseErrors[0].Message)"
    }

    $script:Functions = $ast.FindAll(
        { param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false)
    foreach ($fn in $script:Functions) {
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    $script:ModuleText = Get-Content -LiteralPath $script:ModuleScript -Raw

    # Row shapes mirror local_user_host_list.csv: every value is a string,
    # as Import-Csv hands them over.
    function New-HostRow {
        param(
            [string]$Pc,
            [string]$User,
            [string]$Password = 'Pass1234',
            [string]$Enabled = '1',
            [string]$AutoLogon,
            [switch]$NoAutoLogonColumn
        )
        $row = [ordered]@{
            Enabled   = $Enabled
            NewPCName = $Pc
            UserName  = $User
            Password  = $Password
            Group     = 'Users'
        }
        if (-not $NoAutoLogonColumn) { $row['AutoLogon'] = $AutoLogon }
        return [PSCustomObject]$row
    }
}

Describe 'Get-HostAutoLogonTarget' {

    Context 'Candidate filtering' {

        It 'picks the only enabled user of the selected PC' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password 'P01!')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error    | Should -BeNullOrEmpty
            $r.UserName | Should -Be 'pc01_local'
            $r.Password | Should -Be 'P01!'
        }

        It 'ignores rows of other PCs even when they carry AutoLogon=1' {
            # The normal way to fill the CSV is one flagged row per PC, so a
            # flag on another host must never reach this PC's decision.
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local'),
                (New-HostRow -Pc 'NEW-PC-02' -User 'pc02_admin' -AutoLogon '1'),
                (New-HostRow -Pc 'NEW-PC-03' -User 'pc03_admin' -AutoLogon '1')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error    | Should -BeNullOrEmpty
            $r.UserName | Should -Be 'pc01_local'
        }

        It 'ignores disabled rows' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'old_user' -Enabled '0' -AutoLogon '1'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error    | Should -BeNullOrEmpty
            $r.UserName | Should -Be 'pc01_local'
        }

        It 'errors when the PC has no enabled row' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-02' -User 'pc02_local'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Enabled '0')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Match "NEW-PC-01"
        }

        It 'errors when no host is selected' {
            # Without SELECTED_NEW_PCNAME there is no PC to resolve against,
            # and picking "the first row" would autologon an arbitrary user.
            $rows = @((New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local'))
            foreach ($pc in @('', '   ', $null)) {
                $r = Get-HostAutoLogonTarget -Rows $rows -PcName $pc
                $r.UserName | Should -BeNullOrEmpty
                $r.Error    | Should -Match 'SELECTED_NEW_PCNAME'
            }
        }

        It 'errors when the row set is empty' {
            $r = Get-HostAutoLogonTarget -Rows @() -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Several users on one PC' {

        It 'errors when two users are enabled and none is flagged' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_admin'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_user')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Match 'AutoLogon=1'
        }

        It 'picks the flagged user when several are enabled' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_admin' -AutoLogon '0'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_user' -Password 'U5er!' -AutoLogon '1')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error    | Should -BeNullOrEmpty
            $r.UserName | Should -Be 'pc01_user'
            $r.Password | Should -Be 'U5er!'
        }

        It 'errors when two users of the same PC are flagged' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_admin' -AutoLogon '1'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_user' -AutoLogon '1')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Match 'exactly one'
        }

        It 'treats values other than 1 as unflagged' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_admin' -AutoLogon 'yes'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_user' -AutoLogon '1')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -Be 'pc01_user'
        }
    }

    Context 'Duplicate rows for one user' {

        It 'collapses repeated rows that agree on the password' {
            # Segmented CSVs list the same account more than once (one row per
            # Segment); with one Segment active those rows are one user.
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password 'Same1!'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'PC01_Local' -Password 'Same1!')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error    | Should -BeNullOrEmpty
            $r.UserName | Should -Be 'pc01_local'
        }

        It 'honours a flag set on any row of the collapsed user' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_admin'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_user' -Password 'Same1!'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'PC01_USER' -Password 'Same1!' -AutoLogon '1')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error    | Should -BeNullOrEmpty
            $r.UserName | Should -Be 'pc01_user'
        }

        It 'errors when repeated rows disagree on the password' {
            # Which one the live account carries is unknowable here, and the
            # wrong one arms an autologon that fails at the next boot.
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password 'First1!'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password 'Second1!')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Match 'different passwords'
        }

        It 'treats a case-only password difference as a difference' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password 'Pass1!'),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password 'PASS1!')
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error | Should -Match 'different passwords'
        }
    }

    Context 'Old CSVs without the AutoLogon column' {

        It 'resolves a single user when the column is absent entirely' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -NoAutoLogonColumn)
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.Error    | Should -BeNullOrEmpty
            $r.UserName | Should -Be 'pc01_local'
        }

        It 'errors on several users when the column is absent entirely' {
            $rows = @(
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_admin' -NoAutoLogonColumn),
                (New-HostRow -Pc 'NEW-PC-01' -User 'pc01_user' -NoAutoLogonColumn)
            )
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Match 'AutoLogon=1'
        }
    }

    Context 'Unusable passwords' {

        It 'errors on an empty password' {
            $rows = @((New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password ''))
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Match 'empty password'
        }

        It 'errors on a password that is still ENC: ciphertext' {
            # Import-ModuleCsv leaves the ciphertext in place when it cannot
            # decrypt (no master passphrase / wrong one). Writing that to
            # DefaultPassword would arm a logon that fails at the next boot.
            $rows = @((New-HostRow -Pc 'NEW-PC-01' -User 'pc01_local' -Password 'ENC:Zm9vYmFy'))
            $r = Get-HostAutoLogonTarget -Rows $rows -PcName 'NEW-PC-01'
            $r.UserName | Should -BeNullOrEmpty
            $r.Error    | Should -Match 'encrypted'
        }
    }
}

Describe 'Host-bound mode wiring' {

    It 'uses a token that cannot be a local account name' {
        # Brackets are invalid in Windows local account names, so the token
        # can never collide with a real User value in autologon_list.csv.
        $script:ModuleText | Should -Match '\$HOST_USER_TOKEN\s*=\s*"\[HOSTUSER\]"'
    }

    It 'resolves the per-PC user CSV through the profile data overlay' {
        # A raw Join-Path would read the shipping CSV even when the active
        # profile carries its own data set.
        $script:ModuleText | Should -Match 'Resolve-ModuleDataPath'
        $script:ModuleText | Should -Match 'local_user_host_list\.csv'
    }

    It 'does not wrap the Import-ModuleCsv call in @()' {
        # Import-ModuleCsv returns ,@(rows) so a single row stays an array at
        # the call site. Wrapping the CALL in @() re-collapses every row into
        # one element whose properties are arrays - the module then resolved
        # "user1 user2 user3" as one name. Cost one E2E run on the rig VM
        # (2026-09-16); see memory project_ps51_list_return_pitfall.
        # Comment lines are stripped first - the code comment warning about
        # this very trap spells the bad form out.
        $code = ($script:ModuleText -split "`n" | Where-Object { $_ -notmatch '^\s*#' }) -join "`n"
        $code | Should -Match '\$hostRows = Import-ModuleCsv'
        $code | Should -Not -Match '@\(\s*Import-ModuleCsv'
    }

    It 'requires the local account to exist before arming the logon' {
        # Without this guard a profile that runs autologon before
        # local_user_config parks the next boot on the logon screen.
        $script:ModuleText | Should -Match 'Get-LocalUser'
    }

    It 'pins the host-bound logon to the local SAM with Domain "."' {
        # Verified on the AD lab 2026-09-16: "." works on a domain-joined and
        # on a workgroup PC, survives a pending rename (unlike a baked
        # computer name) and stays unambiguous when the customer's domain
        # holds an account of the same name (unlike an empty value).
        $script:ModuleText | Should -Match 'Domain\s*=\s*"\."'
    }
}
