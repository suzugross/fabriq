# Cisco IOS `do <EXEC command>` support.
#
# From within a configuration mode (GlobalConfig / InterfaceConfig /
# ModuleConfig) `do` runs a privileged EXEC command once and returns
# to the SAME mode and context. Only side-effect-free, non-mode-
# mutating EXEC commands are permitted: a `do` that could change the
# shell mode (configure / disable) or tear down the shell (exit) would
# defeat the entire point of `do`, so those are rejected by whitelist.
#
# The leading 'do' token is intercepted at the REPL (exact match only,
# never abbreviated) BEFORE mode-scoped resolution; this helper then
# re-resolves the remaining tokens against the PrivilegedExec
# vocabulary and dispatches via the existing Invoke-PrivilegedExecCommand.

# EXEC commands runnable through `do`. Both leave $State.Mode and its
# context untouched: Invoke-ShowCommand is read-only, and `reload` is
# performative (no reboot, no mode change). configure / disable / exit
# are deliberately excluded because their handlers call Set-ShellMode
# or set ShouldExit.
$script:FabriqIosDoWhitelist = @('show', 'reload')

function Invoke-FabriqIosDoCommand {
    # $ExecTokens are the tokens AFTER the leading 'do'. $State is the
    # live shell state; this function never mutates $State.Mode or its
    # context fields (CurrentInterface, ConfigModuleName, ...).
    param(
        [string[]]$ExecTokens,
        [hashtable]$State
    )

    if (-not $ExecTokens -or $ExecTokens.Count -eq 0) {
        Write-Host "% Incomplete command. Usage: 'do <EXEC command>' (e.g. 'do show running-config')" -ForegroundColor Red
        return
    }

    # Re-resolve against the privileged EXEC vocabulary so abbreviations
    # like 'sh ru' expand to 'show running-config' exactly as they would
    # at the '#' prompt.
    $resolved = Resolve-FabriqIosCommand -Tokens $ExecTokens -Mode 'PrivilegedExec'
    if ($null -eq $resolved) { return }
    if ($resolved.Error) {
        Write-Host ("% {0}" -f $resolved.Error) -ForegroundColor Red
        return
    }

    if ($resolved.Command -notin $script:FabriqIosDoWhitelist) {
        Write-Host ("% '{0}' cannot be run from configuration mode with 'do'. Allowed: {1}." -f `
            $resolved.Command, ($script:FabriqIosDoWhitelist -join ', ')) -ForegroundColor Red
        return
    }

    # Dispatch as a privileged EXEC command. For 'show' / 'reload' this
    # never calls Set-ShellMode, so the caller stays in the current
    # configuration mode with its context intact.
    Invoke-PrivilegedExecCommand -Resolved $resolved -State $State
}
