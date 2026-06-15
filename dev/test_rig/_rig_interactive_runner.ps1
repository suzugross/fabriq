# ========================================
# Fabriq test rig - interactive session runner (session-b)
# ========================================
# Deployed to the VM at C:\fabriq_test\rig\ and invoked by the 'FabriqRigInteractive'
# scheduled task, which runs in the logged-on desktop session (Session 1). This is how
# user32/SPI/credential/display modules get a real interactive window-station that the
# noninteractive WinRM session (Session 0) cannot provide.
#
# Contract: reads request.json, runs the module via Invoke-SafeCommand, writes result.json
# (incl. sessionId to prove it ran interactively).
# ========================================
$ErrorActionPreference = 'Stop'
$rig     = 'C:\fabriq_test\rig'
$reqPath = Join-Path $rig 'request.json'
$resPath = Join-Path $rig 'result.json'
try {
    if (-not (Test-Path $reqPath)) { return }
    $req = Get-Content $reqPath -Raw | ConvertFrom-Json

    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    Set-Location $req.repoPath
    . .\kernel\common.ps1

    $global:AutoPilotMode          = $true
    $global:FabriqMasterPassphrase = [string]$req.passphrase
    $global:FabriqEvidenceBasePath = 'C:\fabriq_test\evidence'
    $global:FabriqEvidenceRootPath = 'C:\fabriq_test\evidence'
    $global:FabriqTranscriptPath   = 'C:\fabriq_test\logs'
    $env:FABRIQ_EVIDENCE_BASE = 'C:\fabriq_test\evidence'
    $env:SELECTED_KANRI_NO = 'TEST-0001'; $env:FABRIQ_WORKER_NAME = 'ci-harness'
    if ($req.selected) { foreach ($p in $req.selected.PSObject.Properties) { Set-Item -Path "Env:SELECTED_$($p.Name)" -Value ([string]$p.Value) } }
    $env:FABRIQ_SEGMENT = [string]$req.segment; $env:SELECTED_SEGMENT = [string]$req.segment

    $modPath = Join-Path $req.moduleDirVM $req.script
    $r = Invoke-SafeCommand -ScriptBlock { & $modPath } -OperationName $req.moduleName 6>$null 4>$null 5>$null 3>$null 2>$null
    $out = [pscustomobject]@{
        status      = $r.Status
        verified    = $r.Verified
        message     = $r.Message
        durationSec = [math]::Round($r.Duration.TotalSeconds, 2)
        sessionId   = (Get-Process -Id $PID).SessionId
        ranAt       = (Get-Date).ToString('o')
    }
    $out | ConvertTo-Json | Set-Content -Path $resPath -Encoding UTF8
}
catch {
    [pscustomobject]@{ status = '(runner threw)'; message = $_.Exception.Message; sessionId = (Get-Process -Id $PID).SessionId } | ConvertTo-Json | Set-Content -Path $resPath -Encoding UTF8
}
