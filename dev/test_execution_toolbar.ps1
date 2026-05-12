# ========================================
# Manual smoke test: Execution Toolbar (Stage 4c-ext, Option 3)
# ========================================
# Stand-alone harness for apps/fabriq_operator/lib/execution_toolbar.ps1.
# Loads common.ps1 + theme + the toolbar lib, opens the toolbar on its
# dedicated runspace, cycles through Idle/Running states, polls for
# Skip flag creation, and auto-closes after ~30 seconds.
#
# Run from fabriq root:
#   powershell.exe -NoProfile -File .\dev\test_execution_toolbar.ps1
#
# Verification checklist for the operator:
#   1. Toolbar appears in the top-right corner of the primary screen.
#   2. Initial label = "Idle", both buttons disabled.
#   3. After ~3s label changes to "test_module_one", buttons enable.
#   4. After ~12s label changes to a longer module name (truncated
#      with ellipsis - confirms AutoEllipsis works).
#   5. After ~22s label returns to "Idle", buttons disable again.
#   6. DRAG TEST: the toolbar header should be smoothly draggable
#      AT ANY TIME during the test - including during the Idle phase
#      (this validates the dedicated runspace's message loop).
#   7. Click [Skip] during a Running phase:
#      - Console logs "Skip requested (effective only for async modules)"
#      - kernel\json\skip_request.flag is created (test harness deletes
#        it after detection so multiple clicks can be tested cleanly)
#   8. Click [Gyotaq] during a Running phase:
#      - Toolbar disappears briefly (~300ms)
#      - PNG written to evidence\gyotaku\ (no toolbar in the image)
#      - Toolbar reappears at the same location
#      - Console logs "Screenshot saved: <filename>"
#   9. Toolbar closes cleanly at ~30s mark.
# ========================================

$ErrorActionPreference = 'Stop'

# Locate fabriq root (this script lives in dev/).
$fabriqRoot = Split-Path $PSScriptRoot -Parent
Set-Location $fabriqRoot

. (Join-Path $fabriqRoot 'kernel\common.ps1')

# Theme must be loaded before the toolbar lib so $script:bgForm /
# $script:bgPanel / New-StyledButton etc. are in scope for both the
# host script (in case it ever uses them) and the toolbar runspace
# (which re-dot-sources theme.ps1 internally).
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
. (Join-Path $fabriqRoot 'apps\fabriq_operator\lib\theme.ps1')
. (Join-Path $fabriqRoot 'apps\fabriq_operator\lib\execution_toolbar.ps1')

Show-Info "Stage 4c-ext toolbar smoke test (dedicated-runspace mode)"
Show-Info "  - Toolbar runs on its own STA Runspace with its own"
Show-Info "    message loop; the host script's main thread does not"
Show-Info "    need to pump DoEvents."
Show-Info "  - Drag should be smooth at all times (Idle / Running)."
Show-Info "  - Auto-closes after 30 seconds."
Write-Host ""

# Resolve skip flag path the same way the toolbar handler does, so
# the harness can poll and delete-on-detect cleanly.
$asyncCfg     = Get-FabriqAsyncConfig
$skipFlagPath = $asyncCfg.SkipFlagPath
if (-not [System.IO.Path]::IsPathRooted($skipFlagPath)) {
    $skipFlagPath = Join-Path $fabriqRoot $skipFlagPath
}
$skipFlagPath = [System.IO.Path]::GetFullPath($skipFlagPath)

if (Test-Path $skipFlagPath) {
    Remove-Item $skipFlagPath -Force -ErrorAction SilentlyContinue
    Show-Info "Cleared stale skip flag from previous run"
}

Show-ExecutionToolbar

$start = Get-Date
$phase = 0

while (((Get-Date) - $start).TotalSeconds -lt 30) {
    Start-Sleep -Milliseconds 200

    if (Test-Path $skipFlagPath) {
        $content = (Get-Content $skipFlagPath -ErrorAction SilentlyContinue) -join ''
        Show-Success "  [HARNESS] Skip flag detected: '$content'"
        Remove-Item $skipFlagPath -Force -ErrorAction SilentlyContinue
    }

    $elapsed = ((Get-Date) - $start).TotalSeconds

    if ($elapsed -ge 3 -and $phase -lt 1) {
        Update-ExecutionToolbar -ExecutionState 'Running' -ModuleName 'test_module_one'
        Show-Info ("[{0:F1}s] State -> Running (test_module_one)" -f $elapsed)
        $phase = 1
    }
    elseif ($elapsed -ge 12 -and $phase -lt 2) {
        Update-ExecutionToolbar -ExecutionState 'Running' -ModuleName 'this_is_a_module_with_a_very_long_name_to_trigger_ellipsis'
        Show-Info ("[{0:F1}s] State -> Running (long_module_name)" -f $elapsed)
        $phase = 2
    }
    elseif ($elapsed -ge 22 -and $phase -lt 3) {
        Update-ExecutionToolbar -ExecutionState 'Idle'
        Show-Info ("[{0:F1}s] State -> Idle" -f $elapsed)
        $phase = 3
    }
}

Hide-ExecutionToolbar
Show-Success "Toolbar test complete"
