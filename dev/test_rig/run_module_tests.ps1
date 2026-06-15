# ========================================
# Fabriq Module Test Rig - descriptor-driven runner (C5 + C3)
# ========================================
# Reads each module's co-located test.psd1 (C2 contract) and, over WinRM against
# an isolated test VM, runs: sync -> route -> fixture -> envelope -> Invoke-SafeCommand ->
# independent oracle -> verdict -> in-guest teardown. Emits a table and (optionally) JSON.
#
# Dev/test tooling: NOT shipped to kitted PCs, not part of the kernel. Pass/fail is
# decided by the deterministic contract+oracle layers; AI is not in the run-time path (C6).
#
# Revert automation is NOT possible here (the physical VMware host is unmanaged; vmrun is
# unreachable from this sibling guest). Instead:
#   cleanup = none     -> leave (idempotent / isolated change)
#   cleanup = undo      -> run in-guest teardown steps to restore state
#   cleanup = snapshot  -> CANNOT auto-revert; reported as "MANUAL REVERT REQUIRED"
#
# Routing (current rig capability):
#   context = interactive   -> SKIP (needs the session-(b) interactive runner; C3 next)
#   winrmSafe = $false      -> SKIP (would drop the mgmt transport; needs console/spare-NIC)
#   else                    -> run over WinRM
#
# Usage:
#   powershell.exe -File dev\test_rig\run_module_tests.ps1 -Password <pw> [-SyncRepo] [-Module a,b] [-JsonReport out.json]
# Exit code: 0 = no FAIL/ERROR verdicts, 1 = at least one FAIL/ERROR.
# ========================================
[CmdletBinding()]
param(
    [string[]]$Module,
    [string]$ComputerName = '10.1.10.8',
    [pscredential]$Credential,
    [string]$User = 'FabriqTest',
    [string]$Password,                       # plaintext fallback for automation; C9 will replace with a secret store
    [string]$RepoPathOnVM = 'C:\fabriq',
    [string]$JsonReport,
    [switch]$SyncRepo,                       # push kernel\common.ps1 + tested module dirs to the VM before running (fixes stale VM)
    [switch]$Idempotency                     # C7: re-apply each module a 2nd time and assert idempotent.secondRun (slower; off by default)
)
$ErrorActionPreference = 'Stop'

if (-not $Credential) {
    if ($env:FABRIQ_TESTVM_PW) { $Credential = New-Object pscredential($User, (ConvertTo-SecureString $env:FABRIQ_TESTVM_PW -AsPlainText -Force)) }
    elseif ($Password)         { $Credential = New-Object pscredential($User, (ConvertTo-SecureString $Password -AsPlainText -Force)) }
    else                       { $Credential = Get-Credential -UserName $User -Message 'Fabriq test VM' }
}

# --- discover descriptors ---
$repoRoot   = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulesDir = Join-Path $repoRoot 'modules'
$allModuleDirs = Get-ChildItem -Path $modulesDir -Directory -Recurse -Depth 1 |
    Where-Object { Test-Path (Join-Path $_.FullName 'module.csv') }
$descriptors = $allModuleDirs | ForEach-Object {
    $p = Join-Path $_.FullName 'test.psd1'
    if (Test-Path $p) { [pscustomobject]@{ Dir = $_.FullName; Desc = $p } }
} | Where-Object { $_ }
if ($Module) {
    # -File arg binding passes "a,b,c" as a single string; normalize to an array
    $Module = @($Module) | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $descriptors = $descriptors | Where-Object { (Split-Path $_.Dir -Leaf) -in $Module }
}

$covered   = $descriptors | ForEach-Object { Split-Path $_.Dir -Leaf }
$uncovered = $allModuleDirs | ForEach-Object { Split-Path $_.FullName -Leaf } | Where-Object { $_ -notin $covered }

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Fabriq Module Test Rig  (VM: $ComputerName)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("Descriptors: {0}   Uncovered: {1}   SyncRepo: {2}" -f $descriptors.Count, $uncovered.Count, [bool]$SyncRepo) -ForegroundColor Gray
Write-Host ""

# --- remote runner (executes ON the VM where the state lives) ---
$remote = {
    param($RepoPath, $ModuleDirVM, $ModuleName, $Scn, $ApplyMode, $Idempotency)
    $res = [ordered]@{ status='(not run)'; verified=$null; message=''; durationSec=0; oracleVerdict='(none)'; oracleDetail=''; teardown=''; via=''; secondStatus='' }
    try {
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        Set-Location $RepoPath
        . .\kernel\common.ps1

        $global:AutoPilotMode          = [bool]$Scn.envelope.autopilot
        $global:FabriqMasterPassphrase = [string]$Scn.envelope.passphrase
        $global:FabriqEvidenceBasePath = 'C:\fabriq_test\evidence'
        $global:FabriqEvidenceRootPath = 'C:\fabriq_test\evidence'
        $global:FabriqTranscriptPath   = 'C:\fabriq_test\logs'
        $env:FABRIQ_EVIDENCE_BASE      = 'C:\fabriq_test\evidence'
        $env:SELECTED_KANRI_NO = 'TEST-0001'; $env:FABRIQ_WORKER_NAME = 'ci-harness'
        if ($Scn.envelope.selected) { foreach ($k in $Scn.envelope.selected.Keys) { Set-Item -Path "Env:SELECTED_$k" -Value ([string]$Scn.envelope.selected[$k]) } }
        $env:FABRIQ_SEGMENT = [string]$Scn.envelope.segment; $env:SELECTED_SEGMENT = [string]$Scn.envelope.segment

        # fixtures
        $fixOk = $true; $fixDetail = ''
        foreach ($fx in @($Scn.fixture)) {
            switch ($fx.type) {
                'create-files' { foreach ($p in $fx.paths) { $d = Split-Path $p -Parent; if ($d -and -not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }; Set-Content -Path $p -Value ($(if ($fx.content) { $fx.content } else { 'fixture' })) -Force } }
                'stage-asset'  {
                    $from = Join-Path $ModuleDirVM $fx.from; $to = Join-Path $ModuleDirVM $fx.to
                    if (-not (Test-Path $from)) { $fixOk = $false; $fixDetail = "missing asset: $from" }
                    else { if ((Test-Path $to) -and -not (Test-Path "$to.harnessbak")) { Copy-Item $to "$to.harnessbak" -Force }; Copy-Item $from $to -Force }
                }
                'create-localuser' { try { if (-not (Get-LocalUser -Name $fx.name -EA SilentlyContinue)) { New-LocalUser -Name $fx.name -Password (ConvertTo-SecureString 'Fixture#123' -AsPlainText -Force) -EA Stop | Out-Null } } catch { $fixOk = $false; $fixDetail = $_.Exception.Message } }
                'expr'    { try { & ([scriptblock]::Create($fx.run)) | Out-Null } catch { $fixOk = $false; $fixDetail = $_.Exception.Message } }
                default   { }
            }
        }
        if (-not $fixOk) { $res.status = '(fixture failed)'; $res.message = $fixDetail; return [pscustomobject]$res }

        # run -- factored so idempotency can apply twice. inproc (Session 0 / WinRM) or task (Session 1 / session-b)
        $modPath = Join-Path $ModuleDirVM $Scn.script
        if (-not (Test-Path $modPath)) { $res.status = '(missing entry)'; return [pscustomobject]$res }
        function Invoke-Apply {
            $o = @{ status='(null result)'; verified=$null; message=''; dur=0; via='' }
            if ($ApplyMode -eq 'task') {
                $req = @{ repoPath=$RepoPath; moduleDirVM=$ModuleDirVM; script=$Scn.script; moduleName=$ModuleName; passphrase=[string]$Scn.envelope.passphrase; selected=$Scn.envelope.selected; segment=[string]$Scn.envelope.segment } | ConvertTo-Json -Compress
                Set-Content 'C:\fabriq_test\rig\request.json' -Value $req -Encoding UTF8
                Remove-Item 'C:\fabriq_test\rig\result.json' -ErrorAction SilentlyContinue
                try { Start-ScheduledTask -TaskName 'FabriqRigInteractive' -ErrorAction Stop } catch { $o.status='(no session-b task)'; $o.message=$_.Exception.Message; return $o }
                $tr = $null; $deadline = (Get-Date).AddSeconds(120)
                while ((Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 800; if (Test-Path 'C:\fabriq_test\rig\result.json') { try { $tr = Get-Content 'C:\fabriq_test\rig\result.json' -Raw | ConvertFrom-Json } catch { $tr = $null }; if ($tr) { break } } }
                if ($tr) { $o.status=$tr.status; $o.verified=$tr.verified; $o.message=$tr.message; $o.dur=$tr.durationSec; $o.via="session-b(s$($tr.sessionId))" } else { $o.status='(task timeout)'; $o.via='session-b' }
            } else {
                $r = Invoke-SafeCommand -ScriptBlock { & $modPath } -OperationName $ModuleName 6>$null 4>$null 5>$null 3>$null 2>$null
                if ($r) { $o.status=$r.Status; $o.verified=$r.Verified; $o.message=$r.Message; $o.dur=[math]::Round($r.Duration.TotalSeconds, 2) } else { $o.status='(null result)' }
                $o.via='winrm(s0)'
            }
            return $o
        }
        $a1 = Invoke-Apply
        $res.status=$a1.status; $res.verified=$a1.verified; $res.message=$a1.message; $res.durationSec=$a1.dur; $res.via=$a1.via

        # oracle
        $o = $Scn.oracle; $ov = '(none)'; $od = ''
        switch ($o.type) {
            'registry-csv' {
                $files = @(Get-ChildItem -Path (Join-Path $ModuleDirVM $o.csv) -File -EA SilentlyContinue)
                if (-not $files) { $ov = 'ERROR'; $od = 'no csv' }
                else {
                    $rows = @(); foreach ($f in $files) { $rows += (Import-Csv $f.FullName -Encoding Default | Where-Object { $_.Enabled -eq '1' }) }
                    if (-not $rows) { $ov = 'ERROR'; $od = 'no enabled rows' }
                    else {
                        $fail = 0; $n = 0
                        foreach ($row in $rows) {
                            $n++
                            try { $a = (Get-ItemProperty -Path ("Registry::" + $row.KeyPath) -Name $row.KeyName -EA Stop).$($row.KeyName) } catch { $a = $null }
                            $m = $false; if ($null -ne $a) { try { $m = ([int64]$a -eq [int64]$row.Value) } catch { $m = ("$a" -eq "$($row.Value)") } }
                            if (-not $m) { $fail++ }
                        }
                        if ($fail -eq 0) { $ov = 'PASS'; $od = "$n/$n" } else { $ov = 'FAIL'; $od = "$fail/$n mismatch" }
                    }
                }
            }
            'file-exists' {
                $mode = if ($o.mode) { $o.mode } else { 'present' }; $bad = 0; $n = 0
                foreach ($p in $o.paths) { $n++; $ex = Test-Path $p; if (($mode -eq 'present' -and -not $ex) -or ($mode -eq 'absent' -and $ex)) { $bad++ } }
                if ($bad -eq 0) { $ov = 'PASS'; $od = "$mode $n/$n" } else { $ov = 'FAIL'; $od = "$mode $bad/$n wrong" }
            }
            'state-query' { try { $val = & ([scriptblock]::Create($o.query)); $exp = $o.expect.value; if ("$val" -eq "$exp") { $ov = 'PASS'; $od = "$val" } else { $ov = 'FAIL'; $od = "got '$val' exp '$exp'" } } catch { $ov = 'ERROR'; $od = $_.Exception.Message } }
            'command'     { try { $out = & ([scriptblock]::Create($o.run)); $ok = $true; if ($null -ne $o.equals) { $ok = ("$out" -eq "$($o.equals)") } elseif ($o.match) { $ok = ("$out" -match $o.match) }; if ($ok) { $ov = 'PASS'; $od = "$out" } else { $ov = 'FAIL'; $od = "$out" } } catch { $ov = 'ERROR'; $od = $_.Exception.Message } }
            'self-verified' { $ov = 'SELF'; $od = "module Verified=$($res.verified)" }
            'none'          { $ov = 'NONE'; $od = [string]$o.reason }
            default         { $ov = '(none)'; $od = "unknown oracle type: $($o.type)" }
        }
        $res.oracleVerdict = $ov; $res.oracleDetail = $od

        # idempotency 2nd run (C7) -- re-apply WITHOUT re-fixture/teardown, capture status to assert against secondRun
        if ($Idempotency -and $null -ne $Scn.idempotent.secondRun -and $res.status -in @('Success', 'Skipped', 'Partial')) {
            $a2 = Invoke-Apply
            $res.secondStatus = $a2.status
        }

        # teardown (in-guest undo) - only when cleanup = undo
        if ($Scn.cleanup -eq 'undo') {
            foreach ($td in @($Scn.teardown)) {
                switch ($td.type) {
                    'delete-files'     { foreach ($p in $td.paths) { if (Test-Path $p) { Remove-Item $p -Recurse -Force -EA SilentlyContinue } } }
                    'delete-localuser' { try { Remove-LocalUser -Name $td.name -EA SilentlyContinue } catch { } }
                    'reg-delete'       { try { Remove-ItemProperty -Path $td.path -Name $td.name -EA SilentlyContinue } catch { } }
                    'restore-asset'    { $t = Join-Path $ModuleDirVM $td.path; if (Test-Path "$t.harnessbak") { Copy-Item "$t.harnessbak" $t -Force; Remove-Item "$t.harnessbak" -Force -EA SilentlyContinue } }
                    'expr'             { try { & ([scriptblock]::Create($td.run)) | Out-Null } catch { } }
                    default            { }
                }
            }
            $res.teardown = 'undone'
        } else { $res.teardown = [string]$Scn.cleanup }
    }
    catch { $res.status = '(threw)'; $res.message = $_.Exception.Message }
    return [pscustomobject]$res
}

# --- run ---
$sess = New-PSSession -ComputerName $ComputerName -Credential $Credential
$results = @()
$manualRevert = @()
try {
    if ($SyncRepo) {
        Write-Host "Syncing repo to VM (kernel\common.ps1 + tested module dirs)..." -ForegroundColor Gray
        Copy-Item -ToSession $sess (Join-Path $repoRoot 'kernel\common.ps1') (Join-Path $RepoPathOnVM 'kernel\common.ps1') -Force
        foreach ($entry in $descriptors) {
            $rel = $entry.Dir.Substring($repoRoot.Length).TrimStart('\')
            $dst = Join-Path $RepoPathOnVM $rel
            Invoke-Command -Session $sess -ScriptBlock { param($d) if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null } } -ArgumentList $dst
            Copy-Item -ToSession $sess "$($entry.Dir)\*" $dst -Recurse -Force
        }
    }

    foreach ($entry in $descriptors) {
        $d = Import-PowerShellDataFile $entry.Desc
        $rel = $entry.Dir.Substring($repoRoot.Length).TrimStart('\')
        $moduleDirVM = Join-Path $RepoPathOnVM $rel
        foreach ($scn in @($d.scenarios)) {
            $row = [ordered]@{ Module=$d.module; Scenario=$scn.name; Verdict=''; Status=''; Verified=''; Idem='-'; Via=''; Oracle=''; Cleanup=''; Dur=''; Detail='' }

            if ($scn.winrmSafe -eq $false) {
                $row.Verdict = 'SKIP'; $row.Detail = 'winrmSafe=false (needs console/spare-NIC path)'
            }
            else {
                $applyMode = if ($scn.context -eq 'interactive') { 'task' } else { 'inproc' }
                $rr = Invoke-Command -Session $sess -ScriptBlock $remote -ArgumentList $RepoPathOnVM, $moduleDirVM, $d.module, $scn, $applyMode, $Idempotency
                $row.Status = $rr.status; $row.Verified = $rr.verified; $row.Dur = $rr.durationSec; $row.Via = $rr.via
                $row.Oracle = "$($rr.oracleVerdict) $($rr.oracleDetail)".Trim()
                $row.Cleanup = $rr.teardown

                $statusOk = @($scn.expect.status) -contains $rr.status
                $vexp = $scn.expect.verified
                $vOk  = ($vexp -eq 'any') -or ("$($rr.verified)" -eq "$vexp")
                if     (-not $statusOk)                          { $row.Verdict = 'FAIL';  $row.Detail = "status '$($rr.status)' not in expected" }
                elseif ($rr.oracleVerdict -eq 'FAIL')            { $row.Verdict = 'FAIL';  $row.Detail = "oracle: $($rr.oracleDetail)" }
                elseif ($rr.oracleVerdict -eq 'ERROR')           { $row.Verdict = 'ERROR'; $row.Detail = "oracle: $($rr.oracleDetail)" }
                elseif (-not $vOk)                               { $row.Verdict = 'FAIL';  $row.Detail = "verified '$($rr.verified)' != '$vexp'" }
                elseif ($rr.oracleVerdict -eq 'PASS')            { $row.Verdict = 'PASS' }
                elseif ($rr.oracleVerdict -in @('SELF','NONE'))  { $row.Verdict = "PASS($($rr.oracleVerdict))"; $row.Detail = $rr.oracleDetail }
                else                                             { $row.Verdict = 'PASS?'; $row.Detail = $rr.oracleDetail }

                if ($rr.message -and -not $row.Detail) { $row.Detail = $rr.message }

                # idempotency verdict (C7): independent of the 1st-run verdict
                if ($Idempotency -and $null -ne $scn.idempotent.secondRun -and $rr.secondStatus) {
                    if ($rr.secondStatus -eq $scn.idempotent.secondRun) { $row.Idem = 'OK' }
                    else {
                        $row.Idem = "FAIL($($rr.secondStatus))"
                        if ($row.Verdict -notlike 'FAIL*' -and $row.Verdict -ne 'ERROR') { $row.Verdict = 'FAIL'; $row.Detail = "idempotency: 2nd run '$($rr.secondStatus)' != expected '$($scn.idempotent.secondRun)'" }
                    }
                }

                if ($scn.cleanup -eq 'snapshot') { $manualRevert += "$($d.module)/$($scn.name)" }
            }
            $results += [pscustomobject]$row
        }
    }
}
finally { Remove-PSSession $sess }

# --- report ---
$results | Format-Table Module, Scenario, Verdict, Status, Idem, Via, Cleanup, Dur, Oracle -AutoSize
if ($manualRevert.Count -gt 0) {
    Write-Host ""
    Write-Host ("MANUAL REVERT REQUIRED (cleanup=snapshot): {0}" -f ($manualRevert -join ', ')) -ForegroundColor Yellow
}
if ($uncovered.Count -gt 0) {
    Write-Host ""
    Write-Host ("UNCOVERED ({0}): {1}" -f $uncovered.Count, ($uncovered -join ', ')) -ForegroundColor DarkYellow
}
if ($JsonReport) {
    $payload = [pscustomobject]@{ computerName=$ComputerName; results=$results; manualRevert=$manualRevert; uncovered=$uncovered }
    $payload | ConvertTo-Json -Depth 6 | Set-Content -Path $JsonReport -Encoding UTF8
    Write-Host "`nJSON report -> $JsonReport" -ForegroundColor Gray
}

$bad = @($results | Where-Object { $_.Verdict -like 'FAIL*' -or $_.Verdict -eq 'ERROR' }).Count
Write-Host ""
Write-Host ("Summary: {0} scenario(s) | FAIL/ERROR: {1} | manual-revert: {2}" -f $results.Count, $bad, $manualRevert.Count) -ForegroundColor $(if ($bad -gt 0) { 'Red' } else { 'Green' })
exit ($(if ($bad -gt 0) { 1 } else { 0 }))
