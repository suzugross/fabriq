# ========================================
# Fabriq Test Rig - scenario runner (Layer 2, 80/20)
# ========================================
# Runs an ORDERED list of module steps under ONE shared envelope on the VM, in sequence,
# WITHOUT teardown between steps (modules stack like a Profile), then checks a combined
# finalOracle, then tears down in reverse. This exercises real modules in profile order on
# real OS -- the integration that single-module tests and the kernel Pester suite don't.
#
# It deliberately does NOT re-run each module's individual oracle (that's run_module_tests.ps1's
# job) and does NOT drive the real Invoke-BatchExecution (that's covered by the kernel Pester
# suite; AST-driving it headless is low-ROI). __RESTART__ resume is out of scope (structurally
# tied to the real Fabriq.exe relaunch).
#
# v1: noninteractive steps only. interactive / winrmSafe=false steps are reported SKIP.
#
# Scenario file (dev/test_rig/scenarios/<name>.psd1):
#   @{ schema=1; name='...'; description='...'
#      envelope=@{ autopilot=$true; passphrase=''; selected=@{}; segment='' }
#      steps=@( @{ module='reg_hklm_config'; scenario='apply' }, ... )
#      finalOracle=@{ type='command'; run='<single-quoted expr>'; equals='True' } }   # or state-query/file-exists
#
# Usage: powershell.exe -File dev\test_rig\run_scenario.ps1 -Scenario <name|path> -Password <pw> [-SyncRepo]
# Exit code: 0 = no step FAIL + finalOracle ok, 1 otherwise.
# ========================================
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Scenario,
    [string]$ComputerName = '10.1.10.8',
    [pscredential]$Credential,
    [string]$User = 'FabriqTest',
    [string]$Password,
    [string]$RepoPathOnVM = 'C:\fabriq',
    [string]$JsonReport,
    [switch]$SyncRepo
)
$ErrorActionPreference = 'Stop'
if (-not $Credential) {
    if ($env:FABRIQ_TESTVM_PW) { $Credential = New-Object pscredential($User, (ConvertTo-SecureString $env:FABRIQ_TESTVM_PW -AsPlainText -Force)) }
    elseif ($Password)         { $Credential = New-Object pscredential($User, (ConvertTo-SecureString $Password -AsPlainText -Force)) }
    else                       { $Credential = Get-Credential -UserName $User -Message 'Fabriq test VM' }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
# resolve scenario file
$scnPath = if (Test-Path $Scenario) { $Scenario } else { Join-Path $PSScriptRoot "scenarios\$Scenario.psd1" }
if (-not (Test-Path $scnPath)) { Write-Host "[ERROR] scenario not found: $scnPath" -ForegroundColor Red; exit 1 }
$scn = Import-PowerShellDataFile $scnPath

# locally resolve each step's module descriptor -> script/context/fixture/teardown/expect
$steps = @()
foreach ($st in @($scn.steps)) {
    $dir = Get-ChildItem -Path (Join-Path $repoRoot 'modules') -Directory -Recurse -Depth 1 |
        Where-Object { (Split-Path $_.FullName -Leaf) -eq $st.module -and (Test-Path (Join-Path $_.FullName 'test.psd1')) } | Select-Object -First 1
    if (-not $dir) { Write-Host "[ERROR] step module/descriptor not found: $($st.module)" -ForegroundColor Red; exit 1 }
    $d = Import-PowerShellDataFile (Join-Path $dir.FullName 'test.psd1')
    $sc = @($d.scenarios) | Where-Object { $_.name -eq $st.scenario } | Select-Object -First 1
    if (-not $sc) { Write-Host "[ERROR] scenario '$($st.scenario)' not in $($st.module)/test.psd1" -ForegroundColor Red; exit 1 }
    $rel = $dir.FullName.Substring($repoRoot.Length).TrimStart('\')
    $steps += @{
        module = $st.module; moduleDirVM = (Join-Path $RepoPathOnVM $rel); localDir = $dir.FullName
        script = $sc.script; context = $sc.context; winrmSafe = $sc.winrmSafe
        fixture = $sc.fixture; teardown = $sc.teardown; cleanup = $sc.cleanup
        expectStatus = @($sc.expect.status)
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Scenario: $($scn.name)  (VM: $ComputerName)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("Steps: {0}   SyncRepo: {1}" -f $steps.Count, [bool]$SyncRepo) -ForegroundColor Gray
if ($scn.description) { Write-Host "  $($scn.description)" -ForegroundColor DarkGray }
Write-Host ""

$sess = New-PSSession -ComputerName $ComputerName -Credential $Credential
try {
    if ($SyncRepo) {
        Write-Host "Syncing repo to VM..." -ForegroundColor Gray
        Copy-Item -ToSession $sess (Join-Path $repoRoot 'kernel\common.ps1') (Join-Path $RepoPathOnVM 'kernel\common.ps1') -Force
        foreach ($s in $steps) {
            $rel = $s.localDir.Substring($repoRoot.Length).TrimStart('\'); $dst = Join-Path $RepoPathOnVM $rel
            Invoke-Command -Session $sess -ScriptBlock { param($d) if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null } } -ArgumentList $dst
            Copy-Item -ToSession $sess "$($s.localDir)\*" $dst -Recurse -Force
        }
    }

    $remote = {
        param($RepoPath, $Env, $Steps, $FinalOracle)
        $out = @{ steps = @(); final = @{ verdict = '-'; detail = '' } }
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        Set-Location $RepoPath
        . .\kernel\common.ps1
        # shared envelope
        $global:AutoPilotMode = [bool]$Env.autopilot
        $global:FabriqMasterPassphrase = [string]$Env.passphrase
        $global:FabriqEvidenceBasePath = 'C:\fabriq_test\evidence'; $global:FabriqEvidenceRootPath = 'C:\fabriq_test\evidence'; $global:FabriqTranscriptPath = 'C:\fabriq_test\logs'
        $env:FABRIQ_EVIDENCE_BASE = 'C:\fabriq_test\evidence'; $env:SELECTED_KANRI_NO = 'TEST-0001'; $env:FABRIQ_WORKER_NAME = 'ci-harness'
        if ($Env.selected) { foreach ($k in $Env.selected.Keys) { Set-Item -Path "Env:SELECTED_$k" -Value ([string]$Env.selected[$k]) } }
        $env:FABRIQ_SEGMENT = [string]$Env.segment; $env:SELECTED_SEGMENT = [string]$Env.segment

        # run each step in order (no teardown between)
        foreach ($s in $Steps) {
            $r = @{ module = $s.module; status = '(skip)'; detail = '' }
            if ($s.context -eq 'interactive' -or $s.winrmSafe -eq $false) {
                $r.status = '(skip)'; $r.detail = 'v1: noninteractive only'
            } else {
                foreach ($fx in @($s.fixture)) {
                    switch ($fx.type) {
                        'create-files' { foreach ($p in $fx.paths) { $dd = Split-Path $p -Parent; if ($dd -and -not (Test-Path $dd)) { New-Item -ItemType Directory -Force -Path $dd | Out-Null }; Set-Content -Path $p -Value 'fixture' -Force } }
                        'stage-asset'  { $from = Join-Path $s.moduleDirVM $fx.from; $to = Join-Path $s.moduleDirVM $fx.to; if (Test-Path $from) { if ((Test-Path $to) -and -not (Test-Path "$to.harnessbak")) { Copy-Item $to "$to.harnessbak" -Force }; Copy-Item $from $to -Force } }
                        'expr'         { try { & ([scriptblock]::Create($fx.run)) | Out-Null } catch { } }
                        default        { }
                    }
                }
                $mod = Join-Path $s.moduleDirVM $s.script
                if (-not (Test-Path $mod)) { $r.status = '(missing entry)' }
                else { $rr = Invoke-SafeCommand -ScriptBlock { & $mod } -OperationName $s.module 6>$null 4>$null 5>$null 3>$null 2>$null; $r.status = if ($rr) { $rr.status } else { '(null)' } }
            }
            $out.steps += $r
        }

        # combined finalOracle
        if ($FinalOracle -and $FinalOracle.type) {
            $ov = '(none)'; $od = ''
            switch ($FinalOracle.type) {
                'command'     { try { $o = & ([scriptblock]::Create($FinalOracle.run)); $ok = $true; if ($null -ne $FinalOracle.equals) { $ok = ("$o" -eq "$($FinalOracle.equals)") } elseif ($FinalOracle.match) { $ok = ("$o" -match $FinalOracle.match) }; $ov = if ($ok) { 'PASS' } else { 'FAIL' }; $od = "$o" } catch { $ov = 'ERROR'; $od = $_.Exception.Message } }
                'state-query' { try { $v = & ([scriptblock]::Create($FinalOracle.query)); $ov = if ("$v" -eq "$($FinalOracle.expect.value)") { 'PASS' } else { 'FAIL' }; $od = "$v" } catch { $ov = 'ERROR'; $od = $_.Exception.Message } }
                'file-exists' { $mode = if ($FinalOracle.mode) { $FinalOracle.mode } else { 'present' }; $bad = 0; foreach ($p in $FinalOracle.paths) { $ex = Test-Path $p; if (($mode -eq 'present' -and -not $ex) -or ($mode -eq 'absent' -and $ex)) { $bad++ } }; $ov = if ($bad -eq 0) { 'PASS' } else { 'FAIL' }; $od = "$mode" }
                default       { $ov = '(none)'; $od = "unknown type $($FinalOracle.type)" }
            }
            $out.final = @{ verdict = $ov; detail = $od }
        }

        # teardown in reverse (cleanup=undo only)
        for ($i = $Steps.Count - 1; $i -ge 0; $i--) {
            $s = $Steps[$i]
            if ($s.cleanup -eq 'undo') {
                foreach ($td in @($s.teardown)) {
                    switch ($td.type) {
                        'delete-files'     { foreach ($p in $td.paths) { if (Test-Path $p) { Remove-Item $p -Recurse -Force -EA SilentlyContinue } } }
                        'delete-localuser' { try { Remove-LocalUser -Name $td.name -EA SilentlyContinue } catch { } }
                        'reg-delete'       { try { Remove-ItemProperty -Path $td.path -Name $td.name -EA SilentlyContinue } catch { } }
                        'restore-asset'    { $t = Join-Path $s.moduleDirVM $td.path; if (Test-Path "$t.harnessbak") { Copy-Item "$t.harnessbak" $t -Force; Remove-Item "$t.harnessbak" -Force -EA SilentlyContinue } }
                        'expr'             { try { & ([scriptblock]::Create($td.run)) | Out-Null } catch { } }
                        default            { }
                    }
                }
            }
        }
        return [pscustomobject]$out
    }

    $rr = Invoke-Command -Session $sess -ScriptBlock $remote -ArgumentList $RepoPathOnVM, $scn.envelope, $steps, $scn.finalOracle
}
finally { Remove-PSSession $sess }

# verdict
$rows = @()
$fail = 0
for ($i = 0; $i -lt $steps.Count; $i++) {
    $st = $steps[$i]; $res = $rr.steps[$i]
    $verdict = if ($res.status -eq '(skip)') { 'SKIP' } elseif (@($st.expectStatus) -contains $res.status) { 'PASS' } else { 'FAIL'; }
    if ($verdict -eq 'FAIL') { $fail++ }
    $rows += [pscustomobject]@{ '#' = ($i + 1); Module = $st.module; Status = $res.status; Verdict = $verdict; Detail = $res.detail }
}
$rows | Format-Table '#', Module, Status, Verdict, Detail -AutoSize
Write-Host ("finalOracle: {0} {1}" -f $rr.final.verdict, $rr.final.detail) -ForegroundColor $(if ($rr.final.verdict -eq 'PASS' -or $rr.final.verdict -eq '-') { 'Green' } else { 'Red' })
if ($rr.final.verdict -eq 'FAIL' -or $rr.final.verdict -eq 'ERROR') { $fail++ }

if ($JsonReport) { [pscustomobject]@{ scenario = $scn.name; steps = $rows; final = $rr.final } | ConvertTo-Json -Depth 6 | Set-Content $JsonReport -Encoding UTF8 }
Write-Host ""
Write-Host ("Summary: {0} step(s) | FAIL: {1}" -f $steps.Count, $fail) -ForegroundColor $(if ($fail -gt 0) { 'Red' } else { 'Green' })
exit ($(if ($fail -gt 0) { 1 } else { 0 }))
