# ============================================================
# Fabriq IOS - Cisco IOS-style command shell over Fabriq.
# Entry point with self-spawning subprocess isolation.
# Comments are English-only per project policy.
# ============================================================

# Self-spawn guard. When invoked in-process (e.g. from
# fabriq_operator's FabriqApps dialog via & $appPath), re-launch
# in an isolated powershell.exe subprocess and return. This keeps
# any PSReadLine key handlers, environment-variable mutations,
# and global-scope state confined to the child process.
if (-not $env:FABRIQ_IOS_SUBPROCESS) {
    $env:FABRIQ_IOS_SUBPROCESS = '1'
    try {
        $self = $PSCommandPath
        Start-Process powershell.exe `
            -ArgumentList @(
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', "`"$self`""
            ) `
            -Wait
    } finally {
        Remove-Item Env:FABRIQ_IOS_SUBPROCESS -ErrorAction SilentlyContinue
    }
    return
}

# --- The following runs only inside the isolated subprocess. ---

$ErrorActionPreference = 'Stop'
$script:FabriqIosRoot = $PSScriptRoot
$script:FabriqRoot    = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path

# Load Fabriq kernel common library (Show-Info, Import-ModuleCsv,
# Test-MasterPassphrase, Unprotect-FabriqValue, etc.).
. (Join-Path $script:FabriqRoot 'kernel\common.ps1')

# Defensive log-output suppression. fabriq_ios is a sub-project that
# never participates in the parent fabriq's audit trail (no execution
# history, no on-disk evidence capture, no transcript). The dispatch
# path (lib/dispatch.ps1) deliberately does not call any of these
# functions, but a module might. Shadow them with no-ops at the
# global scope so any accidental invocation is silently dropped.
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

# Read fabriq_ios's independent VERSION file (semver, separate from
# kernel). Used by `show version` and `show running-config`.
$script:FabriqIosVersion = '0.0.0'
$_verFile = Join-Path $PSScriptRoot 'VERSION'
if (Test-Path $_verFile) {
    $script:FabriqIosVersion = (Get-Content -Path $_verFile -Raw).Trim()
}

# Suppress per-module Confirm-ModuleExecution prompts: in fabriq_ios
# the act of typing the command IS the confirmation (Cisco IOS UX).
# Wait-KeyPress between modules is also unwanted in a REPL.
$global:AutoPilotMode    = $true
$global:AutoPilotWaitSec = 0

# Load fabriq_ios libraries.
. (Join-Path $PSScriptRoot 'lib\shell_state.ps1')
. (Join-Path $PSScriptRoot 'lib\parser.ps1')
. (Join-Path $PSScriptRoot 'lib\prompt.ps1')
. (Join-Path $PSScriptRoot 'lib\syslog.ps1')
. (Join-Path $PSScriptRoot 'lib\help.ps1')
. (Join-Path $PSScriptRoot 'lib\dispatch.ps1')
. (Join-Path $PSScriptRoot 'lib\completer.ps1')
. (Join-Path $PSScriptRoot 'lib\modes\user_exec.ps1')
. (Join-Path $PSScriptRoot 'lib\modes\privileged_exec.ps1')
. (Join-Path $PSScriptRoot 'lib\modes\global_config.ps1')
. (Join-Path $PSScriptRoot 'lib\modes\interface_config.ps1')
. (Join-Path $PSScriptRoot 'lib\modes\module_config.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\show.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\hostname.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\interface.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\ip_address.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\categories.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\module.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\enable_disable.ps1')
. (Join-Path $PSScriptRoot 'lib\commands\do.ps1')


function Initialize-FabriqIos {
    # Validate PSReadLine availability. fabriq_ios uses
    # [Microsoft.PowerShell.PSConsoleReadLine]::ReadLine(...) so the
    # module must be loaded before Start-FabriqIosShell runs.
    if (-not (Get-Module -Name PSReadLine)) {
        try {
            Import-Module PSReadLine -ErrorAction Stop
        } catch {
            throw "PSReadLine is required for fabriq_ios but is not available: $($_.Exception.Message)"
        }
    }

    # Disable history-based prediction inside the subprocess. The
    # shell ships its own Cisco-style Tab and `?` completion; PSReadLine
    # prediction (InlineView / ListView, default since PSReadLine 2.2)
    # would render suggested completions in the same screen area we
    # use for our manual candidate rows, causing visual overlap. The
    # Set-PSReadLineOption -PredictionSource parameter is unavailable
    # on PSReadLine < 2.1 - swallow the error there.
    try {
        Set-PSReadLineOption -PredictionSource None -ErrorAction Stop
    } catch {}
}

function Show-Banner {
    $bannerPath = Join-Path $PSScriptRoot 'data\version_banner.txt'
    if (Test-Path $bannerPath) {
        Get-Content -Path $bannerPath -Raw -Encoding UTF8 | Write-Host
    }
}

function Start-FabriqIosShell {
    $state = New-ShellState

    Register-FabriqIosCompleter -State $state

    # Override the prompt so PSConsoleReadLine.ReadLine renders our
    # mode-aware Cisco-style prompt automatically. Subprocess-scoped
    # so the parent shell is never affected.
    Set-Item Function:\global:prompt {
        $s = $global:_FabriqIosShellState
        if ($s) {
            return ((Get-FabriqIosPrompt -State $s) + ' ')
        }
        return 'fabriq> '
    }

    while (-not $state.ShouldExit) {
        # Render the prompt manually before each ReadLine call.
        # PSConsoleReadLine.ReadLine() does NOT invoke the prompt
        # function when called programmatically - that is normally
        # the PowerShell Console host's responsibility, but our
        # subprocess REPL bypasses the host's main loop. The
        # global:prompt override above still matters for Tab and ?
        # handlers that call [...PSConsoleReadLine]::InvokePrompt().
        $promptText = Get-FabriqIosPrompt -State $state
        [Console]::Write(('{0} ' -f $promptText))

        $line = $null
        try {
            $line = [Microsoft.PowerShell.PSConsoleReadLine]::ReadLine($host.Runspace, $ExecutionContext)
        } catch {
            Write-Host ""
            Write-Host ("% Read error: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $state.ShouldExit = $true
            break
        }

        # `?` was pressed at an EMPTY prompt: the handler in
        # lib/completer.ps1 deferred to the REPL via AcceptLine
        # because the full mode-level help is too long to redraw
        # cleanly under InvokePrompt. Non-empty buffer cases are now
        # handled inline by the chord handler (Cisco-style: show
        # candidates, restore buffer, continue editing).
        if ($global:_FabriqIosPendingHelpRequest) {
            $global:_FabriqIosPendingHelpRequest = $false
            $global:_FabriqIosPendingHelpBuffer  = $null

            Write-Host ""
            Show-FabriqIosHelp -Mode $state.Mode
            continue
        }

        if ($null -eq $line) { continue }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $tokens = ConvertTo-FabriqIosTokens $line
        if (-not $tokens -or $tokens.Count -eq 0) { continue }

        # Special: 'help' typed at the prompt also shows mode help.
        if ($tokens[0] -ieq 'help') {
            Show-FabriqIosHelp -Mode $state.Mode
            continue
        }

        # Cisco IOS `do <EXEC command>`: from a configuration mode, run a
        # privileged EXEC command (show / reload) without leaving the
        # current mode. Exact 'do' match only - it is never abbreviated,
        # so it cannot shadow a real command (no config verb starts 'do')
        # and 'do' is intentionally absent from the resolver vocabulary.
        if (($tokens[0] -ieq 'do') -and
            ($state.Mode -in @('GlobalConfig', 'InterfaceConfig', 'ModuleConfig'))) {
            $execTokens = @()
            if ($tokens.Count -gt 1) {
                $execTokens = @($tokens[1..($tokens.Count - 1)])
            }
            try {
                Invoke-FabriqIosDoCommand -ExecTokens $execTokens -State $state
            } catch {
                Write-Host ("% {0}" -f $_.Exception.Message) -ForegroundColor Red
            }
            continue
        }

        $resolved = Resolve-FabriqIosCommand -Tokens $tokens -Mode $state.Mode
        if ($null -eq $resolved) { continue }
        if ($resolved.Error) {
            Write-Host ("% {0}" -f $resolved.Error) -ForegroundColor Red
            continue
        }

        try {
            switch ($state.Mode) {
                'UserExec'        { Invoke-UserExecCommand        -Resolved $resolved -State $state }
                'PrivilegedExec'  { Invoke-PrivilegedExecCommand  -Resolved $resolved -State $state }
                'GlobalConfig'    { Invoke-GlobalConfigCommand    -Resolved $resolved -State $state }
                'InterfaceConfig' { Invoke-InterfaceConfigCommand -Resolved $resolved -State $state }
                'ModuleConfig'    { Invoke-ModuleConfigCommand    -Resolved $resolved -State $state }
                default {
                    Write-Host "% Unknown shell mode: $($state.Mode)" -ForegroundColor Red
                }
            }
        } catch {
            Write-Host ("% {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
}


# Entry orchestration.
try {
    Initialize-FabriqIos
    Show-Banner
    Start-FabriqIosShell
} catch {
    Write-Host ""
    Write-Host "Fabriq IOS terminated: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Press Enter to exit." -ForegroundColor DarkGray
    [void](Read-Host)
}
