# ============================================================
# Fabriq BackUper - Backup/Restore satellite over Fabriq.
# Entry point with self-spawning subprocess isolation.
# Comments are English-only per project policy.
# ============================================================

# Self-spawn guard. When invoked in-process (e.g. via & $appPath),
# re-launch in an isolated powershell.exe subprocess and return.
# This keeps PSReadLine key handlers, env-var mutations, and
# global-scope state confined to the child process.
#
# -NoNewWindow makes the child reuse the parent's console window so
# we don't end up with two visible conhost windows when launched from
# Fabriq_BackUper.exe (the C# launcher already creates one fresh
# console via UseShellExecute = true). Process isolation is preserved;
# only the console window is shared.
if (-not $env:FABRIQ_BACKUPER_SUBPROCESS) {
    $env:FABRIQ_BACKUPER_SUBPROCESS = '1'
    try {
        $self = $PSCommandPath
        Start-Process powershell.exe `
            -ArgumentList @(
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', "`"$self`""
            ) `
            -Wait `
            -NoNewWindow
    } finally {
        Remove-Item Env:FABRIQ_BACKUPER_SUBPROCESS -ErrorAction SilentlyContinue
    }
    return
}

# --- The following runs only inside the isolated subprocess. ---

$ErrorActionPreference = 'Stop'
$script:FabriqBackuperRoot = $PSScriptRoot
$script:FabriqRoot         = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path

# Pre-load .NET assemblies needed by the WinForms UI BEFORE dot-sourcing
# any UI library (theme.ps1 also self-loads these defensively).
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
}
catch {
    Write-Host "[FATAL] Failed to load WinForms / Drawing assemblies: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "        .NET Framework 4.x is required." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    return
}

# Load Fabriq kernel common library.
try {
    . (Join-Path $script:FabriqRoot 'kernel\common.ps1')
}
catch {
    Write-Host "[FATAL] Failed to dot-source kernel\common.ps1: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "        FabriqRoot resolved as: $script:FabriqRoot" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    return
}

# Defensive log-output suppression. fabriq_backuper is a satellite
# that never participates in the parent fabriq's audit trail (no
# execution history, no on-disk evidence capture, no transcript).
# A wrapped module might call these — shadow them with no-ops so
# any accidental invocation is silently dropped.
foreach ($_logFn in @(
    'Initialize-ExecutionHistory',
    'Restore-ExecutionHistory',
    'Write-ExecutionHistory',
    'Add-ExecutionResult',
    'Export-ExecutionHistory',
    'Export-HtmlChecklist',
    'Initialize-EvidenceBasePath',
    'Capture-ScreenEvidence'
)) {
    if (Get-Item "Function:$_logFn" -ErrorAction SilentlyContinue) {
        Set-Item "Function:Global:$_logFn" -Value { } -Force
    }
}

# Read independent VERSION file.
$script:BackuperVersion = '0.0.0'
$_verFile = Join-Path $PSScriptRoot 'VERSION'
if (Test-Path $_verFile) {
    $script:BackuperVersion = (Get-Content -Path $_verFile -Raw).Trim()
}

# Suppress per-module Confirm-ModuleExecution prompts: the FabriqBackUper
# UI is the canonical confirmation surface; wrapped modules should run
# without their own Y/N prompts.
$global:AutoPilotMode    = $true
$global:AutoPilotWaitSec = 0

# Load FabriqBackUper libraries.
$libsToLoad = @(
    'lib\hostlist_reader.ps1',
    'lib\manifest_aggregator.ps1',
    'lib\ui\console_menu.ps1',     # legacy console UI, kept as fallback
    'lib\engine.ps1',
    'lib\ui\theme.ps1',
    'lib\ui\csv_io.ps1',           # Phase 2.7
    'lib\ui\user_selector.ps1',    # Phase 2.7
    'lib\ui\userdata_edit_dialog.ps1', # Phase 2.7
    'lib\ui\unc_helper.ps1',
    'lib\ui\unc_connect_dialog.ps1',
    'lib\ui\mode_select_view.ps1',
    'lib\ui\backup_view.ps1',
    'lib\ui\restore_view.ps1',
    'lib\ui\progress_view.ps1',
    'lib\ui\main_form.ps1'
)
foreach ($rel in $libsToLoad) {
    $abs = Join-Path $PSScriptRoot $rel
    try {
        . $abs
    }
    catch {
        Write-Host "[FATAL] Failed to dot-source: $rel" -ForegroundColor Red
        Write-Host "        $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "        $($_.ScriptStackTrace)" -ForegroundColor DarkGray
        Read-Host "Press Enter to exit"
        return
    }
}

# ============================================================
# Welcome banner
# ============================================================
Clear-Host
Write-Host ""
Show-Separator
Write-Host "  Fabriq BackUper  v$($script:BackuperVersion)" -ForegroundColor Cyan
Write-Host "  Backup/Restore satellite over Fabriq kernel $((Get-Content (Join-Path $script:FabriqRoot 'kernel\KERNEL_VERSION') -Raw).Trim())" -ForegroundColor DarkGray
Show-Separator
Write-Host ""

# ============================================================
# Step 1: Admin check (robocopy /B, registry HKLM write require admin)
# ============================================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required."
    Show-Info  "Please re-launch Fabriq_BackUper.exe with admin rights."
    Write-Host ""
    Read-Host "Press Enter to exit"
    return
}

# ============================================================
# Step 2: Ensure master passphrase is set (inherit if available)
# ============================================================
if ([string]::IsNullOrWhiteSpace($global:FabriqMasterPassphrase)) {
    Show-Info "Master passphrase required to read hostlist."
    $verifyPath = Join-Path $script:FabriqRoot 'kernel\txt\passphrase_verify.txt'
    if (-not (Test-Path $verifyPath)) {
        Show-Error "Passphrase verify token not found: $verifyPath"
        Show-Error "FabriqBackUper requires fabriq to be initialized first."
        Read-Host "Press Enter to exit"
        return
    }
    $attempts = 0
    while ($attempts -lt 3) {
        $secure = Read-Host -Prompt 'Fabriq Master Passphrase' -AsSecureString
        $plain  = [System.Net.NetworkCredential]::new('', $secure).Password
        if (Test-MasterPassphrase -Passphrase $plain -VerifyTokenPath $verifyPath) {
            $global:FabriqMasterPassphrase = $plain
            Show-Success "Passphrase accepted."
            break
        }
        $attempts++
        Show-Error "Invalid passphrase. Attempt $attempts of 3."
    }
    if ([string]::IsNullOrWhiteSpace($global:FabriqMasterPassphrase)) {
        Show-Error "Maximum attempts exceeded."
        Read-Host "Press Enter to exit"
        return
    }
} else {
    Show-Info "Inheriting master passphrase from parent fabriq session."
}

Write-Host ""

# ============================================================
# Step 3: Launch WinForms GUI (Phase 2.1)
# Legacy console menu (Show-MainMenu / Invoke-BackuperEngine)
# is still loaded above for fallback; GUI is the canonical path.
# ============================================================
try {
    Start-FabriqBackuperGui `
        -BackuperVersion $script:BackuperVersion `
        -BackuperRoot    $script:FabriqBackuperRoot `
        -FabriqRoot      $script:FabriqRoot
}
catch {
    Show-Error "GUI launch failed: $($_.Exception.Message)"
    Show-Error $_.ScriptStackTrace
    Read-Host "Press Enter to exit"
}

Write-Host ""
Show-Info "Fabriq BackUper session ended."
Write-Host ""
