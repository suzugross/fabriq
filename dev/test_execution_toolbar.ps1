# ========================================
# Manual smoke test: Execution Toolbar
# ========================================
# Stand-alone harness for apps/fabriq_operator/lib/execution_toolbar.ps1.
# Loads common.ps1 + theme + the toolbar lib, opens the toolbar on its
# dedicated runspace, cycles through Idle / Running states + a sequence
# of host-info pushes, polls for Skip flag creation, and auto-closes
# after ~60 seconds.
#
# Run from fabriq root:
#   powershell.exe -NoProfile -File .\dev\test_execution_toolbar.ps1
#
# Verification checklist for the operator:
#   1. Toolbar appears in the top-right corner of the primary screen.
#   2. Initial state = "(no host selected)" placeholder, "Status: Idle",
#      both buttons disabled.
#   3. ~3s : host A pushed (hostname + Eth IP/Sub/GW + 2 DNS + 1 printer).
#            Sections appear; rows show real-time match (current vs
#            target) with OK / ! icons and amber-tinted mismatch rows.
#   4. ~8s : module starts -> label changes to "test_module_one",
#            Skip / Gyotaq enable.
#   5. ~16s: host B pushed (different hostname + Wifi instead of Eth
#            + 4 DNS + 3 printers). Toolbar height grows; sections
#            re-render to match new target shape.
#   6. ~26s: long module name -> truncates with ellipsis.
#   7. ~38s: Idle again; buttons disable.
#   8. ~50s: host info cleared -> placeholder returns.
#   9. DRAG TEST anytime: header should be smoothly draggable at all
#      phases (including the Idle phases and the ShowDialog-equivalent
#      blocking on Sleep here).
#  10. Skip click during Running phase -> kernel\json\skip_request.flag
#      created (test harness deletes it on detect).
#  11. Gyotaq click during Running phase -> PNG saved to
#      evidence\gyotaku\, toolbar reappears at same location.
#  12. Toolbar closes cleanly at ~60s mark.
# ========================================

$ErrorActionPreference = 'Stop'

# Locate fabriq root (this script lives in dev/).
$fabriqRoot = Split-Path $PSScriptRoot -Parent
Set-Location $fabriqRoot

. (Join-Path $fabriqRoot 'kernel\common.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
. (Join-Path $fabriqRoot 'apps\fabriq_operator\lib\theme.ps1')
. (Join-Path $fabriqRoot 'apps\fabriq_operator\lib\execution_toolbar.ps1')

Show-Info "Execution toolbar smoke test (host-info panel + state cycle)"
Show-Info "  - Drag should be smooth at all times."
Show-Info "  - Auto-closes after 60 seconds."
Write-Host ""

# Resolve skip flag path so the harness can poll and delete-on-detect.
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

# Sample host fixtures. Real values from $env:COMPUTERNAME and the
# local network are unlikely to match these targets, so rows will
# render as mismatches - which is what we want to visually confirm
# the amber-tint / OK / ! visualisation.
$hostA = @{
    Hostname    = 'KIT-A12'
    KanriNo     = 'admin01'
    Pin         = '1234'
    EthIP       = '192.168.1.10'
    EthSubnet   = '24'
    EthGateway  = '192.168.1.1'
    WifiIP      = ''
    WifiSubnet  = ''
    WifiGateway = ''
    DNS         = @('8.8.8.8', '1.1.1.1')
    Printers    = @(
        @{ Name = 'HP_LaserJet_M404'; Driver = 'HP LaserJet Pro M404'; Port = 'IP_192.168.1.50' }
    )
}

$hostB = @{
    Hostname    = 'KIT-B07'
    KanriNo     = 'admin02'
    Pin         = ''
    EthIP       = ''
    EthSubnet   = ''
    EthGateway  = ''
    WifiIP      = '10.0.0.5'
    WifiSubnet  = '24'
    WifiGateway = '10.0.0.1'
    DNS         = @('192.168.10.1', '192.168.10.2', '8.8.8.8', '1.1.1.1')
    Printers    = @(
        @{ Name = 'HP_LaserJet_M404';   Driver = 'HP LaserJet Pro M404';   Port = 'IP_10.0.0.50' },
        @{ Name = 'Canon_iR-ADV_C5750'; Driver = 'Canon Generic Plus PCL6'; Port = 'IP_10.0.0.51' },
        @{ Name = 'Epson_LP-S6160';     Driver = 'EPSON LP-S6160';           Port = 'IP_10.0.0.52' }
    )
}

Show-ExecutionToolbar

$start = Get-Date
$phase = 0

while (((Get-Date) - $start).TotalSeconds -lt 60) {
    Start-Sleep -Milliseconds 200

    if (Test-Path $skipFlagPath) {
        $content = (Get-Content $skipFlagPath -ErrorAction SilentlyContinue) -join ''
        Show-Success "  [HARNESS] Skip flag detected: '$content'"
        Remove-Item $skipFlagPath -Force -ErrorAction SilentlyContinue
    }

    $elapsed = ((Get-Date) - $start).TotalSeconds

    if ($elapsed -ge 3 -and $phase -lt 1) {
        Update-ExecutionToolbar -ExecutionState 'Idle' -TargetHostInfo $hostA
        Show-Info ("[{0:F1}s] Host -> A (hostname + Eth + 2 DNS + 1 printer)" -f $elapsed)
        $phase = 1
    }
    elseif ($elapsed -ge 8 -and $phase -lt 2) {
        Update-ExecutionToolbar -ExecutionState 'Running' -ModuleName 'test_module_one'
        Show-Info ("[{0:F1}s] State -> Running (test_module_one)" -f $elapsed)
        $phase = 2
    }
    elseif ($elapsed -ge 16 -and $phase -lt 3) {
        Update-ExecutionToolbar -ExecutionState 'Running' -ModuleName 'test_module_two' -TargetHostInfo $hostB
        Show-Info ("[{0:F1}s] Host -> B (Wifi + 4 DNS + 3 printers)" -f $elapsed)
        $phase = 3
    }
    elseif ($elapsed -ge 26 -and $phase -lt 4) {
        Update-ExecutionToolbar -ExecutionState 'Running' -ModuleName 'this_is_a_module_with_a_very_long_name_to_trigger_ellipsis'
        Show-Info ("[{0:F1}s] Module name -> very long (ellipsis test)" -f $elapsed)
        $phase = 4
    }
    elseif ($elapsed -ge 38 -and $phase -lt 5) {
        Update-ExecutionToolbar -ExecutionState 'Idle'
        Show-Info ("[{0:F1}s] State -> Idle" -f $elapsed)
        $phase = 5
    }
    elseif ($elapsed -ge 50 -and $phase -lt 6) {
        Update-ExecutionToolbar -ExecutionState 'Idle' -TargetHostInfo @{}
        Show-Info ("[{0:F1}s] Host -> empty (placeholder)" -f $elapsed)
        $phase = 6
    }
}

Hide-ExecutionToolbar
Show-Success "Toolbar test complete"
