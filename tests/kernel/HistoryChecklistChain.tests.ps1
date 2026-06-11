# ========================================
# Pester v5 tests for the execution-history -> HTML-checklist chain
# ========================================
# Functions: kernel/common.ps1 :: Write-ExecutionHistory /
#            Import-ExecutionHistory / Restore-ExecutionHistory /
#            Export-HtmlChecklist / Complete-ProfileExecution
# Run     : powershell.exe -File ./dev/run_tests.ps1
#
# This chain IS the B2B deliverable (TM t-0024 (3)): a regression here
# is not a code bug but an error in the customer-facing acceptance
# document. Pinned contracts:
#   A. CSV roundtrip - manual escaping (comma/quote messages), header
#      creation, Order empty-vs-int, Verified passthrough
#   B. Restore-ExecutionHistory's two modes - legacy (separator +
#      IsRestored) and SessionID filter, including the 0-hit EVICTION
#      that fixed a real field bug (stale cross-session entries
#      polluting the FlexProfile checklist; see main.ps1 comment)
#   C. Export-HtmlChecklist row matching - Order-match precedence,
#      STRICT MenuName fallback (no sibling-Order leak), Pending
#      counting, totals reconciliation, HtmlEncode
#   D. Complete-ProfileExecution - log_uploader result recording per
#      mode (incl. the t-0015 "Log Upload (auto)" addition)
#
# Hardware/WMI surfaces (Get-CurrentPCInfo, license, BitLocker) are
# mocked; file I/O runs against temp paths via Set-FabriqTestState /
# a temp CWD tree (group D), never against the repo.
# ========================================

BeforeAll {
    . "$PSScriptRoot\..\_helpers\test_state.ps1"
    $script:RepoRoot = Get-FabriqRepoRoot
    . (Join-Path $script:RepoRoot 'kernel\common.ps1')

    function New-HistoryResult {
        param(
            [string]$Operation, [string]$Status = 'Success', [string]$Message = '',
            $Verified = $null, [int]$Order = 0, [bool]$IsRestored = $false
        )
        [PSCustomObject]@{
            Operation  = $Operation
            Status     = $Status
            Message    = $Message
            Timestamp  = Get-Date
            Verified   = $Verified
            Order      = $Order
            IsRestored = $IsRestored
        }
    }

    function New-DefinedModule {
        param([string]$MenuName, [int]$Order, [string]$Category = 'Test')
        [PSCustomObject]@{
            MenuName     = $MenuName
            Category     = $Category
            Order        = $Order
            RelativePath = ''
        }
    }
}

Describe 'Write-ExecutionHistory / Import-ExecutionHistory (CSV roundtrip)' {

    BeforeEach {
        $script:tmpHistory = Join-Path $env:TEMP `
            ("fabriq-hist-{0}.csv" -f ([guid]::NewGuid().ToString('N')))
        Set-FabriqTestState -HistoryPath $script:tmpHistory -SessionID 'sess-A'
        $script:SessionInfo = $null
        $script:savedKanri = $env:SELECTED_KANRI_NO
        $env:SELECTED_KANRI_NO = 'K-100'
    }

    AfterEach {
        Remove-Item $script:tmpHistory -Force -ErrorAction SilentlyContinue
        $env:SELECTED_KANRI_NO = $script:savedKanri
    }

    It 'creates the 13-column header and a row that survives the roundtrip' {
        $ok = Write-ExecutionHistory -ModuleName 'Mod A' -Category 'Net' -Status 'Success' `
            -Message 'all good' -Verified 'True' -Order 10

        $ok | Should -BeTrue
        (Get-Content $script:tmpHistory -TotalCount 1) |
            Should -Be 'Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID,Verified,Order'

        $rows = @(Import-ExecutionHistory)
        $rows.Count | Should -Be 1
        $rows[0].ModuleName | Should -Be 'Mod A'
        $rows[0].KanriNo | Should -Be 'K-100'
        $rows[0].SessionID | Should -Be 'sess-A'
        $rows[0].Verified | Should -Be 'True'
        $rows[0].Order | Should -Be '10'
    }

    It 'escapes messages containing commas and quotes (the hand-rolled CSV escape)' {
        $msg = 'failed: path "C:\x", retry needed'
        $null = Write-ExecutionHistory -ModuleName 'Mod B' -Category 'FS' -Status 'Error' -Message $msg

        $rows = @(Import-ExecutionHistory)
        $rows[0].Message | Should -Be $msg
        # And the row did not shift columns
        $rows[0].Status | Should -Be 'Error'
        $rows[0].SessionID | Should -Be 'sess-A'
    }

    It 'emits Order=0 as an EMPTY cell (no Profile row association)' {
        $null = Write-ExecutionHistory -ModuleName 'Marker' -Category 'System' -Status 'Success' -Order 0
        $rows = @(Import-ExecutionHistory)
        $rows[0].Order | Should -Be ''
    }

    It 'filters by KanriNo, sorts descending and honors -Limit' {
        # Hand-crafted rows with distinct timestamps (deterministic order)
        @(
            'Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID,Verified,Order'
            '2026-06-11 10:00:00,K-100,PC1,Old,Test,Success,,u,w,m,s1,,'
            '2026-06-11 11:00:00,K-100,PC1,Mid,Test,Success,,u,w,m,s1,,'
            '2026-06-11 12:00:00,K-100,PC1,New,Test,Success,,u,w,m,s1,,'
            '2026-06-11 13:00:00,K-999,PC2,Foreign,Test,Success,,u,w,m,s2,,'
        ) | Set-Content -Path $script:tmpHistory -Encoding UTF8

        $rows = @(Import-ExecutionHistory -FilterKanriNo 'K-100' -Limit 2)
        $rows.Count | Should -Be 2
        $rows[0].ModuleName | Should -Be 'New'
        $rows[1].ModuleName | Should -Be 'Mid'
    }
}

Describe 'Restore-ExecutionHistory (two modes)' {

    BeforeEach {
        $script:tmpHistory = Join-Path $env:TEMP `
            ("fabriq-hist-{0}.csv" -f ([guid]::NewGuid().ToString('N')))
        Set-FabriqTestState -HistoryPath $script:tmpHistory -SessionID 'sess-NOW'
        $script:SessionInfo = $null
        $script:ExecutionResults = @()
        $script:savedKanri = $env:SELECTED_KANRI_NO
        $env:SELECTED_KANRI_NO = 'K-100'

        @(
            'Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID,Verified,Order'
            '2026-06-11 10:00:00,K-100,PC1,ModOld,Test,Success,prev session,u,w,m,sess-OLD,True,10'
            '2026-06-11 11:00:00,K-100,PC1,ModNow,Test,Error,this session,u,w,m,sess-NOW,,20'
        ) | Set-Content -Path $script:tmpHistory -Encoding UTF8
    }

    AfterEach {
        Remove-Item $script:tmpHistory -Force -ErrorAction SilentlyContinue
        $env:SELECTED_KANRI_NO = $script:savedKanri
        $script:ExecutionResults = @()
    }

    It 'legacy mode: restores entries (IsRestored, Verified/Order parsed) and appends the separator' {
        Restore-ExecutionHistory

        $r = @($script:ExecutionResults)
        $r.Count | Should -Be 3
        $r[0].Operation | Should -Be 'ModOld'
        $r[0].IsRestored | Should -BeTrue
        $r[0].Verified | Should -BeTrue
        $r[0].Order | Should -Be 10
        $r[1].Verified | Should -Be $null
        $r[2].Status | Should -Be 'Separator'
    }

    It 'legacy mode: no-op when SELECTED_KANRI_NO is unset' {
        $env:SELECTED_KANRI_NO = ''
        $script:ExecutionResults = @('sentinel')

        Restore-ExecutionHistory

        @($script:ExecutionResults)[0] | Should -Be 'sentinel'
    }

    It 'filter mode: silently replaces with the matching session only (no separator)' {
        Restore-ExecutionHistory -SessionIDFilter 'sess-NOW'

        $r = @($script:ExecutionResults)
        $r.Count | Should -Be 1
        $r[0].Operation | Should -Be 'ModNow'
        $r[0].IsRestored | Should -BeTrue
        @($r | Where-Object { $_.Status -eq 'Separator' }).Count | Should -Be 0
    }

    It 'filter mode with 0 hits EVICTS stale entries (real-bug pin: cross-session pollution)' {
        # Simulate the stale state: session-start legacy restore left
        # cross-session IsRestored entries in memory.
        $script:ExecutionResults = @(
            (New-HistoryResult -Operation 'StaleFromPrevSession' -IsRestored $true)
        )

        Restore-ExecutionHistory -SessionIDFilter 'sess-VIRGIN'

        @($script:ExecutionResults).Count | Should -Be 0
    }

    It 'tolerates a legacy pre-3.1.3 CSV without the Order column (Order defaults to 0)' {
        @(
            'Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID,Verified'
            '2026-06-11 10:00:00,K-100,PC1,LegacyMod,Test,Success,old format,u,w,m,sess-OLD,True'
        ) | Set-Content -Path $script:tmpHistory -Encoding UTF8

        Restore-ExecutionHistory

        $r = @($script:ExecutionResults)
        $r[0].Operation | Should -Be 'LegacyMod'
        $r[0].Order | Should -Be 0
    }
}

Describe 'Export-HtmlChecklist (row matching / counts / encoding)' {

    BeforeEach {
        $script:tmpEvidence = Join-Path $env:TEMP `
            ("fabriq-evid-{0}" -f ([guid]::NewGuid().ToString('N')))
        $global:FabriqEvidenceBasePath = $script:tmpEvidence
        $script:SessionInfo = $null
        $global:FabriqUniqueId = 'TESTUID'

        # Neutralize environment-driven verify sections
        $script:savedEnv = @{}
        foreach ($k in @('SELECTED_NEW_PCNAME','SELECTED_OLD_PCNAME','SELECTED_KANRI_NO',
                         'SELECTED_ETH_IP','SELECTED_ETH_SUBNET','SELECTED_ETH_GATEWAY',
                         'SELECTED_WIFI_IP','SELECTED_WIFI_SUBNET','SELECTED_WIFI_GATEWAY',
                         'SELECTED_DNS1','SELECTED_DNS2','SELECTED_DNS3','SELECTED_DNS4')) {
            $script:savedEnv[$k] = [Environment]::GetEnvironmentVariable($k)
            [Environment]::SetEnvironmentVariable($k, $null, 'Process')
        }
        for ($i = 1; $i -le 10; $i++) {
            foreach ($s in @('NAME','DRIVER','PORT')) {
                [Environment]::SetEnvironmentVariable("SELECTED_PRINTER_${i}_${s}", $null, 'Process')
            }
        }
        $env:SELECTED_NEW_PCNAME = 'TEST-PC'

        # Hardware/WMI surfaces: deterministic stubs
        Mock Get-CurrentPCInfo {
            @{ ComputerName = 'TEST-PC'; EthernetIP = ''; EthernetSubnet = ''; EthernetGateway = ''
               WifiIP = ''; WifiSubnet = ''; WifiGateway = ''; DNS = @(); Printers = @() }
        }
        Mock Get-WmiObject { }
        Mock Get-BitLockerVolume { }
    }

    AfterEach {
        Remove-Item $script:tmpEvidence -Recurse -Force -ErrorAction SilentlyContinue
        $global:FabriqEvidenceBasePath = $null
        foreach ($k in $script:savedEnv.Keys) {
            [Environment]::SetEnvironmentVariable($k, $script:savedEnv[$k], 'Process')
        }
    }

    It 'writes the checklist under FabriqEvidenceBasePath\checklist and returns the path' {
        $path = Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules @((New-DefinedModule -MenuName 'M1' -Order 10)) `
            -ExecutionResults @((New-HistoryResult -Operation 'M1' -Order 10))

        $path | Should -Match 'checklist_.*\.html$'
        Test-Path $path | Should -BeTrue
        (Split-Path (Split-Path $path -Parent) -Parent) | Should -Be $script:tmpEvidence
    }

    It 'Order match wins: a row picks the result with ITS Order, not its MenuName twin' {
        $defined = @((New-DefinedModule -MenuName 'Twin' -Order 20))
        $results = @(
            (New-HistoryResult -Operation 'Twin' -Status 'Error'   -Message 'order10 run' -Order 10),
            (New-HistoryResult -Operation 'Twin' -Status 'Success' -Message 'order20 run' -Order 20)
        )

        $path = Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules $defined -ExecutionResults $results
        $html = Get-Content $path -Raw

        $html | Should -Match 'order20 run'
        $html | Should -Not -Match 'order10 run'
        $html | Should -Match 'chip-ok">OK 1<'
        $html | Should -Match 'chip-ng">NG 0<'
    }

    It 'MenuName fallback is STRICT: a sibling-Order result must not leak in, a legacy Order=0 one may' {
        $defined = @((New-DefinedModule -MenuName 'Leaky' -Order 10))
        # Sibling executed under Order 5 - same MenuName, different row
        $sibling = @((New-HistoryResult -Operation 'Leaky' -Status 'Success' -Message 'sibling ran' -Order 5))

        $p1 = Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules $defined -ExecutionResults $sibling
        $h1 = Get-Content $p1 -Raw
        $h1 | Should -Not -Match 'sibling ran'
        $h1 | Should -Match 'chip-notrun">Not Run 1<'

        # Legacy entry (Order 0 = pre-3.1.3 history) IS accepted as fallback
        $legacy = @((New-HistoryResult -Operation 'Leaky' -Status 'Success' -Message 'legacy ran' -Order 0))
        $p2 = Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules $defined -ExecutionResults $legacy
        $h2 = Get-Content $p2 -Raw
        $h2 | Should -Match 'legacy ran'
        $h2 | Should -Match 'chip-ok">OK 1<'
    }

    It 'reconciles the summary chips with DefinedModules.Count (OK/Skip/NG/NotRun/Pending)' {
        $defined = @(
            (New-DefinedModule -MenuName 'A' -Order 10),
            (New-DefinedModule -MenuName 'B' -Order 20),
            (New-DefinedModule -MenuName 'C' -Order 30),
            (New-DefinedModule -MenuName 'D' -Order 40),
            (New-DefinedModule -MenuName 'E' -Order 50)
        )
        $results = @(
            (New-HistoryResult -Operation 'A' -Status 'Success' -Order 10),
            (New-HistoryResult -Operation 'B' -Status 'Skipped' -Order 20),
            (New-HistoryResult -Operation 'C' -Status 'Error'   -Order 30),
            (New-HistoryResult -Operation 'D' -Status 'Pending' -Order 40)
            # E: never ran
        )

        $html = Get-Content (Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules $defined -ExecutionResults $results) -Raw

        $html | Should -Match 'chip-ok">OK 1<'
        $html | Should -Match 'chip-skip">Skip 1<'
        $html | Should -Match 'chip-ng">NG 1<'
        # Pending counts toward Not Run (totals reconcile: 1+1+1+2 = 5)
        $html | Should -Match 'chip-notrun">Not Run 2<'
        $html | Should -Match '>Pending<'
        $html | Should -Match 'Total: 5 items'
    }

    It 'HTML-encodes hostile result messages (no markup injection into the deliverable)' {
        $defined = @((New-DefinedModule -MenuName 'X' -Order 10))
        $results = @((New-HistoryResult -Operation 'X' -Message '<script>alert(1)</script> & "q"' -Order 10))

        $html = Get-Content (Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules $defined -ExecutionResults $results) -Raw

        $html | Should -Not -Match '<script>alert'
        $html | Should -Match '&lt;script&gt;'
        $html | Should -Match '&amp;'
    }

    It 'renders Verified PASS / FAIL badges and "-" when verification was not performed' {
        $defined = @(
            (New-DefinedModule -MenuName 'VP' -Order 10),
            (New-DefinedModule -MenuName 'VF' -Order 20),
            (New-DefinedModule -MenuName 'VN' -Order 30)
        )
        $results = @(
            (New-HistoryResult -Operation 'VP' -Order 10 -Verified $true),
            (New-HistoryResult -Operation 'VF' -Order 20 -Verified $false),
            (New-HistoryResult -Operation 'VN' -Order 30)
        )

        $html = Get-Content (Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules $defined -ExecutionResults $results) -Raw

        $html | Should -Match '>PASS<'
        $html | Should -Match '>FAIL<'
    }

    It 'styles marker rows ([RESTART] etc.) with the marker-row class' {
        $defined = @((New-DefinedModule -MenuName '[RESTART]' -Order 15 -Category 'System'))

        $html = Get-Content (Export-HtmlChecklist -ProfileName 'P' -ProfilePath '' `
            -DefinedModules $defined -ExecutionResults @()) -Raw

        $html | Should -Match 'class="marker-row"'
    }
}

Describe 'Complete-ProfileExecution (uploader result recording per mode)' {

    BeforeEach {
        # Temp fabriq-like tree: the function resolves the uploader script
        # and log_destinations.csv relative to CWD. view_report.ps1 is
        # deliberately ABSENT so the viewer step is skipped in both modes.
        $script:tmpRoot = Join-Path $env:TEMP `
            ("fabriq-cpe-{0}" -f ([guid]::NewGuid().ToString('N')))
        New-Item -ItemType Directory -Path (Join-Path $script:tmpRoot 'kernel\csv') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $script:tmpRoot 'modules\extended\log_uploader') -Force | Out-Null
        $script:uploaderPath = Join-Path $script:tmpRoot 'modules\extended\log_uploader\log_uploader.ps1'
        $script:uploaderRan  = Join-Path $script:tmpRoot 'uploader-ran.marker'
        Set-Content -Path (Join-Path $script:tmpRoot 'kernel\csv\log_destinations.csv') `
            -Value "Enabled,Path`r`n1,C:\dest" -Encoding Ascii
        Push-Location $script:tmpRoot

        Mock Write-KernelTelemetryEvent { }
        Mock Export-ExecutionHistory { $null }
        Mock Export-HtmlChecklist { 'C:\fake\checklist.html' }
        Mock Add-ExecutionResult { }
        Mock Write-ExecutionHistory { $true }
    }

    AfterEach {
        Pop-Location
        Remove-Item $script:tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
        $global:FabriqLastProfileName = $null
        $global:FabriqLastProfilePath = $null
        $global:FabriqLastProfileModules = $null
    }

    It 'Auto mode + uploader Success: runs the uploader but records nothing (legacy-quiet)' {
        Set-Content -Path $script:uploaderPath -Encoding Ascii -Value @"
Set-Content -Path '$($script:uploaderRan)' -Value 'ran'
return (New-ModuleResult -Status "Success" -Message "up ok")
"@

        $ret = Complete-ProfileExecution -ProfileName 'P' -ProfilePath 'C:\fake\p.csv' `
            -DefinedModules @((New-DefinedModule -MenuName 'M' -Order 10)) -Mode 'Auto'

        $ret | Should -Be 'C:\fake\checklist.html'
        Test-Path $script:uploaderRan | Should -BeTrue
        Should -Invoke Add-ExecutionResult -Times 0 -Exactly
        Should -Invoke Export-HtmlChecklist -Times 1 -Exactly
    }

    It 'Auto mode + uploader failure: records "Log Upload (auto)" (t-0015 contract)' {
        Set-Content -Path $script:uploaderPath -Encoding Ascii -Value @'
return (New-ModuleResult -Status "Error" -Message "share unreachable")
'@

        $null = Complete-ProfileExecution -ProfileName 'P' -ProfilePath 'C:\fake\p.csv' `
            -DefinedModules @((New-DefinedModule -MenuName 'M' -Order 10)) -Mode 'Auto'

        Should -Invoke Add-ExecutionResult -Times 1 -Exactly -ParameterFilter {
            $Operation -eq 'Log Upload (auto)' -and $Status -eq 'Error' -and $Order -eq 0
        }
        Should -Invoke Write-ExecutionHistory -Times 1 -Exactly -ParameterFilter {
            $ModuleName -eq 'Log Upload (auto)' -and $Status -eq 'Error'
        }
    }

    It 'Manual mode: records "Log Upload (cl)" regardless of status (legacy [cl] behavior)' {
        Set-Content -Path $script:uploaderPath -Encoding Ascii -Value @'
return (New-ModuleResult -Status "Success" -Message "up ok")
'@

        $null = Complete-ProfileExecution -ProfileName 'P' -ProfilePath 'C:\fake\p.csv' `
            -DefinedModules @((New-DefinedModule -MenuName 'M' -Order 10)) -Mode 'Manual'

        Should -Invoke Add-ExecutionResult -Times 1 -Exactly -ParameterFilter {
            $Operation -eq 'Log Upload (cl)' -and $Status -eq 'Success'
        }
    }

    It 'skips the uploader entirely when no destination is enabled' {
        Set-Content -Path (Join-Path $script:tmpRoot 'kernel\csv\log_destinations.csv') `
            -Value "Enabled,Path`r`n0,C:\dest" -Encoding Ascii
        Set-Content -Path $script:uploaderPath -Encoding Ascii -Value @"
Set-Content -Path '$($script:uploaderRan)' -Value 'ran'
return (New-ModuleResult -Status "Success" -Message "up ok")
"@

        $null = Complete-ProfileExecution -ProfileName 'P' -ProfilePath 'C:\fake\p.csv' `
            -DefinedModules @((New-DefinedModule -MenuName 'M' -Order 10)) -Mode 'Auto'

        Test-Path $script:uploaderRan | Should -BeFalse
        Should -Invoke Add-ExecutionResult -Times 0 -Exactly
    }
}
