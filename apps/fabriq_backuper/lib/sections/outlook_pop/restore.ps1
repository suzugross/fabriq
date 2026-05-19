# ============================================================
# FabriqBackUper Section: outlook_pop / restore (Phase 2.10.2)
#
# Two-strategy restore with automatic fallback:
#
#   Strategy B (primary, new in 2.10.2):
#     - Import the source profile registry hive (.reg files captured
#       by backup Phase 2.10.1) into the target user's HKCU. Outlook
#       starts up fully configured; the operator only types the
#       password on first send/receive (DPAPI restriction).
#     - Verified empirically 2026-05-16 on s_suzuki -> Administrator
#       and 365 -> 2019 cross-user / cross-version cases.
#
#   Strategy A (fallback, preserved from 2.10.0):
#     - Place PSTs at expected target paths and generate a
#       RESTORE_INSTRUCTIONS.txt cheat sheet for fully-manual operator
#       Outlook wizard setup. Used when Strategy B is not viable
#       (missing reg exports, version mismatch, import failure, or
#       post-import verification mismatch).
#
# Common to both strategies: PST file placement at the target user's
# Documents\<localized-Outlook-Files>\<email>.pst. The userdata
# section is expected to have done the actual file copy; this section
# performs the idempotent rename and verifies presence.
#
# SectionParams (hashtable, all optional):
#   TargetUserProfilePath : profile path of the user whose HKCU to
#                           target. Mirrors the userdata section.
# ============================================================

param(
    [Parameter(Mandatory = $true)][string]$BackuperRoot,
    [Parameter(Mandatory = $true)][string]$FabriqRoot,
    [Parameter(Mandatory = $true)][string]$OldPcName,
    [Parameter(Mandatory = $true)][string]$AggregateBackupDir,
    [hashtable]$SectionParams = @{}
)

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$warnings = @()

# ----------------------------------------------------------
# Strategy B helper functions
# ----------------------------------------------------------

function Get-OutlookInstallInfo {
    # Phase 2.11.0: detect Outlook installation on the local machine.
    # Mirror of the helper in backup.ps1 (identical implementation).
    # Currently used for diagnostic / Summary recording only; no branching
    # decisions are made on the result. Future phases (per-version restore
    # behaviour, e.g. Outlook 2013 PRF fallback) will key off these fields.
    #
    # Returns a hashtable with:
    #   Installed         : $true/$false
    #   RegistryVersion   : '15.0' / '16.0' / $null   (highest version present)
    #   ProductFamily     : 'Outlook 2013' / 'Outlook 2016/2019/365' / $null
    #   InstallType       : 'ClickToRun' / 'MSI' / $null
    #   ProductReleaseIds : raw value from C2R Configuration (or $null)
    #   OutlookExePath    : full path to OUTLOOK.EXE (or $null)
    #   OutlookExeVersion : FileVersion of OUTLOOK.EXE (or $null)
    #   AllVersionsFound  : array of every detected version ('16.0','15.0')
    $result = @{
        Installed         = $false
        RegistryVersion   = $null
        ProductFamily     = $null
        InstallType       = $null
        ProductReleaseIds = $null
        OutlookExePath    = $null
        OutlookExeVersion = $null
        AllVersionsFound  = @()
    }

    $versionsToProbe = @('16.0', '15.0')
    $found = @()
    foreach ($v in $versionsToProbe) {
        $candidateKeys = @(
            "HKLM:\SOFTWARE\Microsoft\Office\$v\Outlook\InstallRoot",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$v\Outlook\InstallRoot"
        )
        $exePath = $null
        foreach ($k in $candidateKeys) {
            if (-not (Test-Path -LiteralPath $k)) { continue }
            try {
                $p = (Get-ItemProperty -LiteralPath $k -ErrorAction Stop).Path
                if ([string]::IsNullOrWhiteSpace($p)) { continue }
                $candidate = Join-Path $p 'OUTLOOK.EXE'
                if (Test-Path -LiteralPath $candidate) {
                    $exePath = $candidate
                    break
                }
            } catch { }
        }
        if ($null -ne $exePath) {
            $found += [PSCustomObject]@{ Version = $v; ExePath = $exePath }
        }
    }

    if ($found.Count -eq 0) { return $result }

    $result.Installed = $true
    $result.AllVersionsFound = @($found | ForEach-Object { $_.Version })

    $primary = $found[0]
    $result.RegistryVersion = $primary.Version
    $result.OutlookExePath  = $primary.ExePath

    try {
        $vi = (Get-Item -LiteralPath $primary.ExePath).VersionInfo
        if ($vi) { $result.OutlookExeVersion = $vi.FileVersion }
    } catch { }

    if ($primary.Version -eq '15.0') {
        $result.ProductFamily = 'Outlook 2013'
    } else {
        $result.ProductFamily = 'Outlook 2016/2019/365'
    }

    $c2rKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path -LiteralPath $c2rKey) {
        try {
            $cfg = Get-ItemProperty -LiteralPath $c2rKey -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace("$($cfg.ProductReleaseIds)")) {
                $result.ProductReleaseIds = "$($cfg.ProductReleaseIds)"
                $result.InstallType = 'ClickToRun'
            }
        } catch { }
    }
    if ($null -eq $result.InstallType) {
        $result.InstallType = 'MSI'
    }

    return $result
}

function Get-RegFileSourceHive {
    # Inspect a .reg file (UTF-16LE) and return the hive prefix used by
    # its first [HIVE\...] header. Examples:
    #   [HKEY_CURRENT_USER\Software\...]          -> 'HKEY_CURRENT_USER'
    #   [HKEY_USERS\S-1-5-21-...\Software\...]    -> 'HKEY_USERS\S-1-5-21-...'
    # Returns $null if no recognisable header is found.
    param([Parameter(Mandatory = $true)][string]$RegPath)
    $text = [System.IO.File]::ReadAllText($RegPath, [System.Text.Encoding]::Unicode)
    $pattern = '\[(HKEY_CURRENT_USER|HKEY_USERS\\[^\\\]]+)\\'
    $m = [System.Text.RegularExpressions.Regex]::Match($text, $pattern)
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Convert-RegFileToTargetHive {
    # Rewrite all '[<SourcePrefix>\...]' headers in a .reg file to use
    # '<TargetPrefix>' instead, producing a new temp .reg with the same
    # UTF-16LE BOM encoding. Returns the temp path. If prefixes match,
    # returns the input path unchanged (no temp file created).
    param(
        [Parameter(Mandatory = $true)][string]$SrcRegPath,
        [Parameter(Mandatory = $true)][string]$SourcePrefix,
        [Parameter(Mandatory = $true)][string]$TargetPrefix
    )
    if ($SourcePrefix -eq $TargetPrefix) { return $SrcRegPath }

    $text = [System.IO.File]::ReadAllText($SrcRegPath, [System.Text.Encoding]::Unicode)
    $escSrc = [System.Text.RegularExpressions.Regex]::Escape($SourcePrefix)
    $rewritten = [System.Text.RegularExpressions.Regex]::Replace(
        $text, "\[$escSrc\\", "[$TargetPrefix\")
    $rewritten = [System.Text.RegularExpressions.Regex]::Replace(
        $rewritten, "\[$escSrc\]", "[$TargetPrefix]")

    $tempBase = [System.IO.Path]::GetTempFileName()
    $tempReg = [System.IO.Path]::ChangeExtension($tempBase, '.reg')
    if (Test-Path -LiteralPath $tempBase) {
        Remove-Item -LiteralPath $tempBase -Force -ErrorAction SilentlyContinue
    }
    [System.IO.File]::WriteAllText($tempReg, $rewritten, [System.Text.Encoding]::Unicode)
    return $tempReg
}

function Invoke-RegImport {
    # Run reg.exe import and capture stdout/stderr + exit code via file
    # redirection. Returns @{ Success; ExitCode; Output }.
    #
    # WHY Start-Process with file redirection instead of '& reg.exe ... 2>&1':
    # reg.exe import writes its success message ("operation completed
    # successfully" / localized variant) to STDERR, not STDOUT. PowerShell's
    # 2>&1 merges stderr into the error stream as ErrorRecord objects;
    # when fabriq's engine sets $ErrorActionPreference = 'Stop' (the
    # default for sections), encountering an ErrorRecord terminates the
    # section even though reg import itself succeeded with exit 0. Using
    # Start-Process with separate file redirects sidesteps the PS stream
    # merge entirely.
    param([Parameter(Mandatory = $true)][string]$RegPath)

    $tempOut = [System.IO.Path]::GetTempFileName()
    $tempErr = [System.IO.Path]::GetTempFileName()
    try {
        $proc = Start-Process -FilePath 'reg.exe' `
            -ArgumentList @('import', $RegPath) `
            -NoNewWindow -Wait -PassThru `
            -RedirectStandardOutput $tempOut `
            -RedirectStandardError  $tempErr
        $exit = $proc.ExitCode
        $stdout = if (Test-Path -LiteralPath $tempOut) { (Get-Content -LiteralPath $tempOut -Raw) } else { '' }
        $stderr = if (Test-Path -LiteralPath $tempErr) { (Get-Content -LiteralPath $tempErr -Raw) } else { '' }
        $combined = (@($stdout, $stderr) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " | "
        return @{
            Success  = ($exit -eq 0)
            ExitCode = $exit
            Output   = $combined.Trim()
        }
    } finally {
        Remove-Item -LiteralPath $tempOut -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempErr -Force -ErrorAction SilentlyContinue
    }
}

function Test-AccountImported {
    # Walk the target hive to verify a specific POP3 account subkey
    # exists with the expected POP3 Server value. Returns
    # @{ Verified=$true/$false; Reason=<string-or-null> }.
    param(
        [Parameter(Mandatory = $true)][string]$HiveDrivePath,
        [Parameter(Mandatory = $true)][string]$OutlookVersion,
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [Parameter(Mandatory = $true)][string]$SubKey,
        [Parameter(Mandatory = $true)][string]$ExpectedPop3Server
    )
    $accountKey = "$HiveDrivePath\Software\Microsoft\Office\$OutlookVersion\Outlook\Profiles\" +
                  "$ProfileName\9375CFF0413111d3B88A00104B2A6676\$SubKey"
    if (-not (Test-Path -LiteralPath $accountKey)) {
        return @{ Verified = $false; Reason = "subkey not found: $accountKey" }
    }
    try {
        $rk = Get-Item -LiteralPath $accountKey -ErrorAction Stop
        $raw = $rk.GetValue('POP3 Server', $null)
        if ($null -eq $raw) {
            return @{ Verified = $false; Reason = 'POP3 Server value missing at imported subkey' }
        }
        $imported = if ($raw -is [byte[]]) {
            [System.Text.Encoding]::Unicode.GetString([byte[]]$raw).TrimEnd([char]0)
        } else { [string]$raw }
        if ($imported -eq $ExpectedPop3Server) {
            return @{ Verified = $true; Reason = $null }
        }
        return @{ Verified = $false; Reason = "POP3 Server mismatch: expected '$ExpectedPop3Server', imported '$imported'" }
    } catch {
        return @{ Verified = $false; Reason = "verify exception: $($_.Exception.Message)" }
    }
}

function Convert-RegFileToStrategyBLight {
    # Phase 2.10.3 / extended in Phase 2.12.0 for cross-version + IMAP.
    #
    # Pre-process a profile .reg file to make Outlook's account-to-store
    # binding cross-PC-portable. Five transforms, all section-aware where
    # needed:
    #
    #   T1 (Phase 2.12.0, cross-version only):
    #     Rewrite '\Office\<srcVer>\Outlook\' -> '\Office\<dstVer>\Outlook\'
    #     in every section header and value path. Applied only when
    #     SourceVersion -ne TargetVersion. No-op for same-version restore.
    #
    #   T2 (Phase 2.10.3 baseline):
    #     Strip "Delivery Store EntryID" / "Delivery Folder EntryID" from
    #     POP/IMAP account subkeys (\9375CFF0...\NNNNNNNN). Scope-limited
    #     so 365's native service-def subkeys are not affected.
    #
    #   T3 (Phase 2.12.0, IMAP cross-PC support):
    #     Strip "IMAP Store EID" from POP/IMAP account subkeys
    #     (\9375CFF0...\NNNNNNNN). IMAP Store EID encodes a binary MAPI
    #     EntryID pointing to the source PC's OST file via per-machine
    #     MAPIUID -> unresolvable on target. Scope-limited same as T2.
    #
    #     IMPORTANT: do NOT strip "Service UID" -- it is load-bearing for
    #     the MAPI service-def loader (Phase 2.12.0 PoC v3 corrupted the
    #     profile this way: Account Settings dialog won't open, dialog
    #     buttons unresponsive, UAC prompts on Profile Management).
    #     Similarly, do NOT strip "Preferences UID" (also present in POP
    #     account subkey, would risk POP regression).
    #
    #   T4 (Phase 2.10.3 baseline):
    #     Rewrite "001f6700" (PR_PST_PATH-like) and "001f0433"
    #     (sharing.xml path) UTF-16LE hex values, replacing source
    #     username with target username. These two value names have no
    #     binary offset tables so are safe to extend.
    #
    #   T5 (Phase 2.12.0, OST cross-PC support):
    #     Drop any subkey whose "001f6700" value decodes to a path ending
    #     in ".ost". These are MAPI service-def subkeys for IMAP message
    #     stores; their OST file references point to source PC paths that
    #     either don't exist or can't be DPAPI-decrypted on the target.
    #     Dropping them lets Outlook auto-create a new OST on first IMAP
    #     sync (the normal IMAP cross-PC behaviour).
    #
    # Returns the path to a new temp .reg with the transforms applied.
    param(
        [Parameter(Mandatory = $true)][string]$SrcRegPath,
        [Parameter(Mandatory = $true)][string]$SourceUserName,
        [Parameter(Mandatory = $true)][string]$TargetUserName,
        [string]$SourceVersion = $null,
        [string]$TargetVersion = $null
    )

    $intAcctRe = '\\9375CFF0413111d3B88A00104B2A6676\\[0-9A-Fa-f]{8}$'

    $content = [System.IO.File]::ReadAllText($SrcRegPath, [System.Text.Encoding]::Unicode)

    # Pre-fix: defensive patch for continuation lines missing trailing
    # '\' (seen in some reg.exe export output edge cases)
    $preLines = $content -split "`r?`n"
    $preFixed = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $preLines.Count; $i++) {
        $cur = $preLines[$i]
        if ($cur -match '[0-9a-f]{2},?$' -and ($i + 1) -lt $preLines.Count) {
            $next = $preLines[$i + 1]
            if ($next -match '^\s+[0-9a-f]{2}' -and $cur -notmatch '\\$') {
                $cur = $cur.TrimEnd() + '\'
            }
        }
        $preFixed.Add($cur)
    }
    $content = $preFixed -join "`r`n"

    # Collapse continuations into single logical lines
    $normalized = $content -replace '\\\r?\n\s+', ''

    # ---- T1: version path rewrite (only when versions differ) ----
    $t1Count = 0
    if (-not [string]::IsNullOrWhiteSpace($SourceVersion) -and
        -not [string]::IsNullOrWhiteSpace($TargetVersion) -and
        $SourceVersion -ne $TargetVersion) {
        $srcVerEsc   = [regex]::Escape($SourceVersion)
        $pathPattern = "\\Office\\$srcVerEsc\\Outlook\\"
        $pathReplace = "\Office\$TargetVersion\Outlook\"
        $t1Count = ([regex]::Matches($normalized, $pathPattern)).Count
        $normalized = $normalized -replace $pathPattern, $pathReplace
    }

    # ---- T5 pre-scan: identify OST-bearing service-def subkeys ----
    $lines = $normalized -split "`r?`n"
    $currentKey = $null
    $ostSections = New-Object System.Collections.Generic.HashSet[string]
    foreach ($line in $lines) {
        if ($line -match '^\[(.+)\]\s*$') {
            $currentKey = $matches[1]
            continue
        }
        if ($null -ne $currentKey -and $line -match '^"001f6700"=hex:(.*)$') {
            $hex = $matches[1] -split ','
            try {
                $bytes = $hex | ForEach-Object { [byte]("0x$_") }
                $s = [System.Text.Encoding]::Unicode.GetString([byte[]]$bytes).TrimEnd([char]0)
                if ($s -match '\.ost$') {
                    [void]$ostSections.Add($currentKey)
                }
            } catch { }
        }
    }

    # Username UTF-16LE hex for T4 rewrite
    $srcUserBytes = [System.Text.Encoding]::Unicode.GetBytes($SourceUserName)
    $dstUserBytes = [System.Text.Encoding]::Unicode.GetBytes($TargetUserName)
    $srcUserHex = ($srcUserBytes | ForEach-Object { '{0:x2}' -f $_ }) -join ','
    $dstUserHex = ($dstUserBytes | ForEach-Object { '{0:x2}' -f $_ }) -join ','

    # ---- Main pass: T2/T3/T4 inline, T5 section drop ----
    $processedLines = New-Object System.Collections.Generic.List[string]
    $currentKey   = $null
    $inAcctSubkey = $false
    $dropSection  = $false
    $stripPop     = 0
    $stripImap    = 0
    $rewritePath  = 0
    $droppedSec   = 0
    foreach ($line in $lines) {
        if ($line -match '^\[(.+)\]\s*$') {
            $currentKey = $matches[1]
            $inAcctSubkey = $currentKey -match $intAcctRe
            if ($ostSections.Contains($currentKey)) {
                $dropSection = $true
                $droppedSec++
                continue
            } else {
                $dropSection = $false
                $processedLines.Add($line)
            }
            continue
        }
        if ($dropSection) { continue }

        if ($inAcctSubkey) {
            if ($line -match '^"Delivery Store EntryID"=' -or
                $line -match '^"Delivery Folder EntryID"=') {
                $stripPop++
                continue
            }
            if ($line -match '^"IMAP Store EID"=') {
                $stripImap++
                continue
            }
        }
        if ($line -match '^"001f6700"=hex:' -or $line -match '^"001f0433"=hex:') {
            $before = $line
            $line = $line -replace [regex]::Escape($srcUserHex), $dstUserHex
            if ($before -ne $line) { $rewritePath++ }
        }
        $processedLines.Add($line)
    }

    Show-Info ("  [BL-transform] T1 version-path=$t1Count  T2 POP-strip=$stripPop  " +
               "T3 IMAP-strip=$stripImap  T4 path-rewrite=$rewritePath  T5 OST-drop=$droppedSec")

    # Re-flow hex value lines back to <=80 cols with '\' continuation
    $outputLines = New-Object System.Collections.Generic.List[string]
    $maxLen = 80
    foreach ($line in $processedLines) {
        if ($line -match '^("[^"]+"=hex:)(.*)$') {
            $header = $matches[1]
            $body   = $matches[2]
            $bytes  = @($body -split ',')
            $singleLine = $header + ($bytes -join ',')
            if ($singleLine.Length -le $maxLen) {
                $outputLines.Add($singleLine)
                continue
            }
            $currentLine = $header
            $indent      = '  '
            $i = 0
            while ($i -lt $bytes.Count) {
                $tok = $bytes[$i]
                $sep = if ($currentLine -eq $header -or $currentLine -eq $indent) { '' } else { ',' }
                $needed = $sep.Length + $tok.Length
                $isLast = ($i -eq ($bytes.Count - 1))
                $reserved = if ($isLast) { 0 } else { 2 }
                if (($currentLine.Length + $needed + $reserved) -le $maxLen) {
                    $currentLine = $currentLine + $sep + $tok
                    $i++
                } else {
                    $outputLines.Add($currentLine + ',\')
                    $currentLine = $indent
                }
            }
            $outputLines.Add($currentLine)
        } else {
            $outputLines.Add($line)
        }
    }

    $finalText = $outputLines -join "`r`n"

    $tempBase = [System.IO.Path]::GetTempFileName()
    $tempReg = [System.IO.Path]::ChangeExtension($tempBase, '.reg')
    if (Test-Path -LiteralPath $tempBase) {
        Remove-Item -LiteralPath $tempBase -Force -ErrorAction SilentlyContinue
    }
    [System.IO.File]::WriteAllText($tempReg, $finalText, [System.Text.Encoding]::Unicode)
    return $tempReg
}

function New-OutlookAccountInfoText {
    # Build a human-readable account-info text used by:
    #   (a) Strategy A fallback -> aggregate/sections/outlook_pop/
    #       RESTORE_INSTRUCTIONS.txt (engineer / operator)
    #   (b) Always-on target-folder copy -> <target_user>\Documents\
    #       <localized_outlook_files>\_account_settings.txt (operator
    #       safety net, travels with the PST file)
    #
    # Content: source/target metadata, per-account email + PST file +
    # POP3/SMTP server settings, and a manual setup procedure that an
    # operator can follow if automatic restore ever has to be redone
    # by hand. All English (feedback_scripts_english_only).
    param(
        [Parameter(Mandatory = $true)]$Manifest,
        [Parameter(Mandatory = $true)][string]$TargetUserProfilePath,
        [Parameter(Mandatory = $true)]$PlannedAccounts,
        [Parameter(Mandatory = $true)]$ResultsByAccount,
        [bool]$StrategyBAttempted = $false,
        [bool]$StrategyBSucceeded = $false,
        [string]$ProfileFilter = $null,
        [bool]$IsCrossVersion = $false,
        [string]$CrossVersionDirection = $null
    )

    # Optional per-profile filter: when writing the target-folder copy
    # we want only the accounts whose PST sits in that profile's folder.
    $effectiveResults = $ResultsByAccount
    if (-not [string]::IsNullOrWhiteSpace($ProfileFilter)) {
        $effectiveResults = @($ResultsByAccount | Where-Object { $_.profile -eq $ProfileFilter })
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('========================================') | Out-Null
    if ($StrategyBSucceeded) {
        $lines.Add(' Outlook POP Account Settings (auto-restored)') | Out-Null
    } else {
        $lines.Add(' Outlook POP Account Restore - Manual Setup Required') | Out-Null
    }
    $lines.Add('========================================') | Out-Null
    $lines.Add('') | Out-Null

    if ($StrategyBAttempted -and -not $StrategyBSucceeded) {
        $lines.Add('NOTE: Strategy B (automatic registry import) was attempted but') | Out-Null
        $lines.Add('did not fully verify. Falling back to manual wizard setup.') | Out-Null
        $lines.Add('See section warnings in the run summary for details.') | Out-Null
        $lines.Add('') | Out-Null
    }
    if ($StrategyBSucceeded) {
        $lines.Add('This file is a reference copy of the account-to-PST mapping and') | Out-Null
        $lines.Add('the full server settings. Automatic restore succeeded; you only') | Out-Null
        $lines.Add('need these settings if you ever have to manually reconfigure') | Out-Null
        $lines.Add('the Outlook account from scratch.') | Out-Null
        $lines.Add('') | Out-Null
    }

    $lines.Add("Source PC      : $($Manifest.computerName)") | Out-Null
    if ($Manifest.sourceUser -and $Manifest.sourceUser.userName) {
        $lines.Add("Source user    : $($Manifest.sourceUser.userName)") | Out-Null
    }
    $tgtUserName = Split-Path $TargetUserProfilePath -Leaf
    $lines.Add("Target user    : $tgtUserName  ($TargetUserProfilePath)") | Out-Null
    $lines.Add("Restore time   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
    $lines.Add("Outlook version: $($Manifest.outlookVersion)") | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($ProfileFilter)) {
        $lines.Add("Profile        : $ProfileFilter") | Out-Null
    }
    $lines.Add('') | Out-Null

    $accountIndex = 0
    foreach ($r in $effectiveResults) {
        $accountIndex++
        $acct = ($PlannedAccounts | Where-Object {
            $_.ProfileName -eq $r.profile -and $_.Account.subKey -eq $r.accountSubKey
        } | Select-Object -First 1).Account

        $lines.Add('----------------------------------------') | Out-Null
        $lines.Add(" Account $accountIndex : $($r.email)   [$($r.status)]") | Out-Null
        $lines.Add('----------------------------------------') | Out-Null

        if ($r.status -ne 'Success') {
            $lines.Add('  ** NOT READY for manual setup **') | Out-Null
            $lines.Add("  Reason: $($r.reason)") | Out-Null
            $lines.Add('') | Out-Null
            continue
        }

        $lines.Add('') | Out-Null
        $lines.Add('  Wizard input (Display Name / Email / Server settings):') | Out-Null
        $lines.Add("    Display name      : $($acct.displayName)") | Out-Null
        $lines.Add("    Email address     : $($acct.email)") | Out-Null
        $lines.Add('    Account type      : POP') | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('  Incoming server (POP3):') | Out-Null
        $lines.Add("    Server            : $($acct.pop3.server)") | Out-Null
        $lines.Add("    Port              : $($acct.pop3.port)") | Out-Null
        $popSsl = if ($acct.pop3.useSSL -eq 1) { 'YES (required)' } else { 'No' }
        $lines.Add("    SSL/TLS           : $popSsl") | Out-Null
        $lines.Add("    Username          : $($acct.pop3.userName)") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('  Outgoing server (SMTP):') | Out-Null
        $lines.Add("    Server            : $($acct.smtp.server)") | Out-Null
        $lines.Add("    Port              : $($acct.smtp.port)") | Out-Null
        $smtpSsl = if ($acct.smtp.useSSL -eq 1) { 'YES (required)' } else { 'No' }
        $lines.Add("    SSL/TLS           : $smtpSsl") | Out-Null
        $smtpAuth = if ($acct.smtp.useAuth -eq 1) { 'YES (required)' } else { 'No' }
        $lines.Add("    Authentication    : $smtpAuth") | Out-Null
        $smtpUser = if ($acct.smtp.userName) { $acct.smtp.userName } else { '(same as POP3)' }
        $lines.Add("    SMTP username     : $smtpUser") | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add('  Existing data file (PST):') | Out-Null
        $lines.Add("    $($r.targetPstPath)") | Out-Null
        $lines.Add('') | Out-Null
    }

    $lines.Add('----------------------------------------') | Out-Null
    $lines.Add(' Manual setup procedure (only if needed):') | Out-Null
    $lines.Add('----------------------------------------') | Out-Null
    $lines.Add('  1. Open Outlook') | Out-Null
    $lines.Add('  2. File > Add Account') | Out-Null
    $lines.Add('  3. Expand "Advanced options" and CHECK "Let me set up my') | Out-Null
    $lines.Add('     account manually" (THIS STEP IS REQUIRED - otherwise') | Out-Null
    $lines.Add('     Outlook will auto-pick IMAP, not POP)') | Out-Null
    $lines.Add('  4. Enter the email address from above, click "Connect"') | Out-Null
    $lines.Add('  5. Choose "POP" as the account type') | Out-Null
    $lines.Add('  6. Enter the server settings exactly as printed above') | Out-Null
    $lines.Add('  7. When asked about data file, choose "Existing Outlook') | Out-Null
    $lines.Add('     Data File" and browse to the PST path printed above') | Out-Null
    $lines.Add('  8. Complete the wizard') | Out-Null
    $lines.Add('  9. Enter the password when Outlook prompts on first') | Out-Null
    $lines.Add('     send/receive (DPAPI restriction: passwords are never') | Out-Null
    $lines.Add('     deployable across machines)') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add(' If the email matches the PST filename (it should, since we') | Out-Null
    $lines.Add(' renamed it to <email>.pst), Outlook attaches the existing') | Out-Null
    $lines.Add(' PST automatically and all old emails / folders / contacts') | Out-Null
    $lines.Add(' will be visible.') | Out-Null
    $lines.Add('') | Out-Null

    if ($IsCrossVersion) {
        $dirLabel = if ($CrossVersionDirection) { " ($CrossVersionDirection)" } else { '' }
        $lines.Add('========================================') | Out-Null
        $lines.Add(" Cross-version restore cleanup steps$dirLabel") | Out-Null
        $lines.Add('========================================') | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add(' The source and target Outlook are different major versions') | Out-Null
        $lines.Add(' (e.g. 2013 -> 2016/2019/365). Outlook on first launch will') | Out-Null
        $lines.Add(' show some prompts that need operator action. These are') | Out-Null
        $lines.Add(' Microsoft behaviours that cannot be suppressed via registry.') | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add(' A. "IMAP Search Folder" warning popup (one-time)') | Out-Null
        $lines.Add('    Outlook tells you old IMAP search folders no longer') | Out-Null
        $lines.Add('    apply to the new OST. Press OK to dismiss; it will') | Out-Null
        $lines.Add('    not appear again.') | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add(' B. An empty "Outlook.pst" gets auto-created') | Out-Null
        $lines.Add('    With cross-version reg-import, the imported profile') | Out-Null
        $lines.Add('    has no "default delivery store" pointer, so Outlook') | Out-Null
        $lines.Add('    creates a fresh empty Outlook.pst and uses it as the') | Out-Null
        $lines.Add('    default. To rebind to the migrated PST:') | Out-Null
        $lines.Add('      1. File > Account Settings > Account Settings') | Out-Null
        $lines.Add('         > Data Files tab') | Out-Null
        $lines.Add('      2. Select the migrated PST') | Out-Null
        $lines.Add('         (<email>.pst under Outlook Files folder)') | Out-Null
        $lines.Add('         and click "Set as Default"') | Out-Null
        $lines.Add('      3. Close Outlook completely, then relaunch') | Out-Null
        $lines.Add('      4. Same Data Files tab: select Outlook.pst and') | Out-Null
        $lines.Add('         click "Remove"') | Out-Null
        $lines.Add('      5. Email tab: select the POP account, click') | Out-Null
        $lines.Add('         "Change Folder", select the Inbox of the') | Out-Null
        $lines.Add('         migrated PST, click OK') | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add(' C. Password input') | Out-Null
        $lines.Add('    On first send/receive, enter the POP and IMAP') | Out-Null
        $lines.Add('    passwords (DPAPI restriction: passwords are never') | Out-Null
        $lines.Add('    portable across machines, this is unavoidable).') | Out-Null
        $lines.Add('') | Out-Null
        $lines.Add(' After these steps, all subsequent launches and send/receive') | Out-Null
        $lines.Add(' will work normally.') | Out-Null
        $lines.Add('') | Out-Null
    }

    $lines.Add('========================================') | Out-Null

    return ($lines -join "`r`n")
}

# ----------------------------------------------------------
# Parse SectionParams
# ----------------------------------------------------------
$targetUserProfilePath = $null
if ($SectionParams.ContainsKey('TargetUserProfilePath') -and `
    -not [string]::IsNullOrWhiteSpace($SectionParams['TargetUserProfilePath'])) {
    $targetUserProfilePath = "$($SectionParams['TargetUserProfilePath'])"
}
if ([string]::IsNullOrWhiteSpace($targetUserProfilePath)) {
    $targetUserProfilePath = $env:USERPROFILE
}

# ----------------------------------------------------------
# Stage 1: read backup manifest, flatten plannedAccounts
# ----------------------------------------------------------
$sectionDir = Join-Path $AggregateBackupDir 'sections\outlook_pop'
$manifestPath = Join-Path $sectionDir 'manifest.json'
if (-not (Test-Path $manifestPath)) {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @("manifest.json not found at: $manifestPath")
    }
}
$manifest = $null
try { $manifest = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json } catch {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @("manifest.json parse error: $($_.Exception.Message)")
    }
}
if ($manifest.manifestType -ne 'fabriq-outlook-pop-backup') {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @("Unexpected manifestType: $($manifest.manifestType)")
    }
}

# ----------------------------------------------------------
# Phase 2.11.0: log source-side install info (from manifest, if backup
# was taken with Phase 2.11.0+ backup.ps1) and probe target-side install.
# Detection only - no branching on these values yet.
# ----------------------------------------------------------
$sourceInstall = $null
if ($manifest.PSObject.Properties.Name -contains 'installedOutlook' -and `
    $null -ne $manifest.installedOutlook) {
    $sourceInstall = $manifest.installedOutlook
    $srcFam  = if ($sourceInstall.productFamily) { $sourceInstall.productFamily } else { 'unknown' }
    $srcType = if ($sourceInstall.installType) { $sourceInstall.installType } else { 'unknown' }
    $srcExe  = if ($sourceInstall.outlookExeVersion) { $sourceInstall.outlookExeVersion } else { 'unknown' }
    Show-Info ("Source Outlook (from manifest): $srcFam (type=$srcType, exeVer=$srcExe)")
} else {
    Show-Info 'Source Outlook (from manifest): not recorded (pre-2.11.0 backup)'
}

$targetInstall = Get-OutlookInstallInfo
if ($targetInstall.Installed) {
    Show-Info ("Target Outlook (this PC): $($targetInstall.ProductFamily) " +
               "(reg=$($targetInstall.RegistryVersion), " +
               "type=$($targetInstall.InstallType), " +
               "exeVer=$($targetInstall.OutlookExeVersion))")
    if ($targetInstall.ProductReleaseIds) {
        Show-Info ("  ProductReleaseIds: $($targetInstall.ProductReleaseIds)")
    }
    if (@($targetInstall.AllVersionsFound).Count -gt 1) {
        Show-Info ("  side-by-side detected: " +
                   ($targetInstall.AllVersionsFound -join ', '))
    }
} else {
    Show-Warning 'Target Outlook (this PC): not detected via HKLM probe'
}

# Cross-family advisory (logged only; no branching). The Strategy B path
# was empirically verified across the 365 -> 2019 case; other cross-family
# combinations may require a future fallback path.
if ($sourceInstall -and $sourceInstall.productFamily -and `
    $targetInstall.Installed -and `
    $sourceInstall.productFamily -ne $targetInstall.ProductFamily) {
    Show-Warning ("Outlook family mismatch: source=$($sourceInstall.productFamily) " +
                  "target=$($targetInstall.ProductFamily) " +
                  "(continuing with current restore path)")
}

# ----------------------------------------------------------
# Phase 2.12.0: cross-version detection.
# Manifest schema 1 records source registry version both at the top level
# (manifest.outlookVersion, present since Phase 2.9.0) and inside
# installedOutlook.registryVersion (Phase 2.11.0+). Target is probed live
# above. When the two registry versions differ, Strategy B-light is
# invoked with version args -> activates T1 (path rewrite) and triggers
# the cross-version operator cleanup section in _account_settings.txt.
# ----------------------------------------------------------
$srcRegVer = "$($manifest.outlookVersion)"
if ($sourceInstall -and $sourceInstall.registryVersion) {
    $srcRegVer = "$($sourceInstall.registryVersion)"
}
$tgtRegVer = $null
if ($targetInstall.Installed) { $tgtRegVer = "$($targetInstall.RegistryVersion)" }

$isCrossVersion = $false
$crossVersionDirection = $null
if (-not [string]::IsNullOrWhiteSpace($srcRegVer) -and
    -not [string]::IsNullOrWhiteSpace($tgtRegVer) -and
    $srcRegVer -ne $tgtRegVer) {
    $isCrossVersion = $true
    $crossVersionDirection = "$srcRegVer -> $tgtRegVer"
    Show-Info ("Cross-version restore detected: $crossVersionDirection " +
               "(Strategy B-light T1 path rewrite will be applied)")
    # 16.0 -> 15.0 direction is unvalidated; 365-only values (Account UID,
    # New Signature, etc.) may confuse 2013. Emit a warning so the operator
    # is aware. Same caveat for 15.0 -> 15.0 (unanticipated equal-version
    # case but logically should reduce to same-version B-light).
    if ($srcRegVer -eq '16.0' -and $tgtRegVer -eq '15.0') {
        Show-Warning ('Cross-version direction 16.0 -> 15.0 is unvalidated. ' +
                      '365-only profile values may not be accepted by Outlook 2013. ' +
                      'Manual operator cleanup or Strategy A fallback may be required.')
    }
}

$plannedAccounts = @()
foreach ($prof in @($manifest.items.profiles)) {
    foreach ($acct in @($prof.accounts)) {
        $plannedAccounts += [PSCustomObject]@{
            ProfileName = $prof.name
            Account     = $acct
        }
    }
}
if ($plannedAccounts.Count -eq 0) {
    return [PSCustomObject]@{
        Status = 'Skipped'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{ note = 'no POP accounts in manifest' }
        Warnings = @($warnings)
    }
}

Show-Info "Preparing $($plannedAccounts.Count) POP account(s) for restore"

# Progress entries: one per POP account
try {
    $uiEntries = foreach ($pa in $plannedAccounts) {
        @{ Id = "$($pa.ProfileName)/$($pa.Account.subKey)"; Label = "$($pa.Account.email)" }
    }
    Initialize-ProgressEntries -Entries @($uiEntries)
} catch { }

$sourceUserProfile = $null
if ($manifest.sourceUser -and $manifest.sourceUser.profilePath) {
    $sourceUserProfile = "$($manifest.sourceUser.profilePath)".TrimEnd('\','/')
}

# ----------------------------------------------------------
# Stage 2: PST placement (common to Strategy A and B)
# Per account: verify PST exists at rebased target path and rename to
# <email>.pst (idempotent). Records per-account result for use by
# subsequent strategy decision.
# ----------------------------------------------------------
$successCount = 0
$failCount    = 0
$resultsByAccount = @()

foreach ($pa in $plannedAccounts) {
    $profileName = $pa.ProfileName
    $a           = $pa.Account
    $entryId     = "$profileName/$($a.subKey)"
    try { Set-EntryStatus -Id $entryId -Status 'InProgress' } catch { }
    Show-Info "[$entryId] $($a.email)"

    $accountResult = [ordered]@{
        profile       = $profileName
        accountSubKey = $a.subKey
        email         = $a.email
        status        = 'Failed'
        reason        = $null
        targetPstPath = $null
        verifyResult  = $null
    }

    if (-not $a.pst -or [string]::IsNullOrWhiteSpace($a.pst.sourcePath)) {
        $msg = "no pst.sourcePath in manifest for $entryId - backup may have failed PST detection."
        $warnings += $msg
        Show-Warning "  $msg"
        $accountResult.reason = 'PST mapping unavailable in manifest'
        $resultsByAccount += $accountResult
        $failCount++
        try { Set-EntryStatus -Id $entryId -Status 'Failed' } catch { }
        continue
    }

    # Cross-user rebase: source\path -> target\path
    $srcPath = "$($a.pst.sourcePath)"
    $rebasedPath = $null
    if ($null -ne $sourceUserProfile -and $srcPath.Length -ge $sourceUserProfile.Length -and `
        $srcPath.Substring(0, $sourceUserProfile.Length) -ieq $sourceUserProfile) {
        $rest = $srcPath.Substring($sourceUserProfile.Length)
        $rebasedPath = "$($targetUserProfilePath.TrimEnd('\','/'))$rest"
    } else {
        $rebasedPath = $srcPath
    }
    Show-Info "  expected PST path at target: $rebasedPath"

    # Target rename path: <folder>\<email>.pst (path-collision-attach)
    $targetDir = Split-Path -Path $rebasedPath -Parent
    $renamedPath = Join-Path $targetDir "$($a.email).pst"

    if (-not (Test-Path -LiteralPath $rebasedPath) -and -not (Test-Path -LiteralPath $renamedPath)) {
        $msg = "PST not at target: '$rebasedPath' and not at '$renamedPath'. Engineer: ensure userdata section includes the Outlook Files folder."
        $warnings += $msg
        Show-Error "  $msg"
        $accountResult.reason = 'PST not found at target'
        $accountResult.targetPstPath = $rebasedPath
        $resultsByAccount += $accountResult
        $failCount++
        try { Set-EntryStatus -Id $entryId -Status 'Failed' } catch { }
        continue
    }

    if ($rebasedPath -ine $renamedPath) {
        if (Test-Path -LiteralPath $renamedPath) {
            Show-Info "  rename target already at <email>.pst (idempotent skip): $renamedPath"
            if (Test-Path -LiteralPath $rebasedPath) {
                $msg = "both '$rebasedPath' and '$renamedPath' exist; will use the renamed one. Engineer: original-name file is orphaned."
                $warnings += $msg
                Show-Warning "  $msg"
            }
        } else {
            try {
                Move-Item -LiteralPath $rebasedPath -Destination $renamedPath -ErrorAction Stop
                Show-Success "  renamed PST -> $renamedPath"
            } catch {
                $msg = "PST rename failed: $($_.Exception.Message). Source: $rebasedPath -> $renamedPath"
                $warnings += $msg
                Show-Error "  $msg"
                $accountResult.reason = "rename failed: $($_.Exception.Message)"
                $accountResult.targetPstPath = $renamedPath
                $resultsByAccount += $accountResult
                $failCount++
                try { Set-EntryStatus -Id $entryId -Status 'Failed' } catch { }
                continue
            }
        }
    }
    $accountResult.targetPstPath = $renamedPath
    $accountResult.status = 'Success'
    $resultsByAccount += $accountResult
    $successCount++
    try { Set-EntryStatus -Id $entryId -Status 'Done' } catch { }
}

# ----------------------------------------------------------
# Stage 3 + 4: Strategy B attempt (reg import + per-account verify)
#
# Viability gate: items.regExports[] must exist and be non-empty.
# Older 2.10.0 backups have no regExports; those go straight to A.
# ----------------------------------------------------------
$regExports = @()
if ($manifest.items.PSObject.Properties.Name -contains 'regExports') {
    $regExports = @($manifest.items.regExports)
}

$strategyBAttempted = $false
$strategyBSucceeded = $false
$strategyBDetails = @()

if ($regExports.Count -gt 0 -and $successCount -gt 0) {
    $strategyBAttempted = $true
    Show-Info ('Strategy B: reg-import path is viable (' + $regExports.Count + ' profile export(s))')

    # Resolve target hive
    $hkcuInfo = Resolve-HkcuRoot
    if ($null -eq $hkcuInfo -or [string]::IsNullOrWhiteSpace($hkcuInfo.PsDrivePath)) {
        $warnings += 'Strategy B: Resolve-HkcuRoot returned null - falling back to Strategy A'
        Show-Warning '  Resolve-HkcuRoot returned null - falling back to Strategy A'
    } else {
        $targetHivePsDrive = $hkcuInfo.PsDrivePath
        $targetHivePrefix  = $hkcuInfo.RegExePath
        if ($hkcuInfo.Redirected) {
            Show-Info "  target hive: $($hkcuInfo.Label) [SID=$($hkcuInfo.SID)]"
        } else {
            Show-Info "  target hive: $($hkcuInfo.Label)"
        }
        Show-Info "  target hive prefix: $targetHivePrefix"

        $outlookVersion = "$($manifest.outlookVersion)"
        if ([string]::IsNullOrWhiteSpace($outlookVersion)) { $outlookVersion = '16.0' }

        # Phase 2.10.3: derive source/target user names for B-light path
        # rewrite. Defaults are safe: if either side is missing the
        # transform's path-rewrite step becomes a no-op (delivery-strip
        # step still runs).
        $blSrcUserName = $null
        if ($manifest.sourceUser -and $manifest.sourceUser.userName) {
            $blSrcUserName = "$($manifest.sourceUser.userName)"
        } elseif ($null -ne $sourceUserProfile) {
            $blSrcUserName = Split-Path -Path $sourceUserProfile -Leaf
        }
        $blDstUserName = Split-Path -Path $targetUserProfilePath -Leaf
        if ([string]::IsNullOrWhiteSpace($blSrcUserName)) { $blSrcUserName = $blDstUserName }
        Show-Info "  [BL-transform] user rewrite: $blSrcUserName -> $blDstUserName"

        $allProfilesVerified = $true
        foreach ($re in $regExports) {
            $profName = "$($re.profileName)"
            $regFile  = "$($re.regFile)"
            $regPath  = Join-Path $sectionDir $regFile

            $perProfile = [ordered]@{
                profileName     = $profName
                regFile         = $regFile
                blTransformed   = $false
                hiveRewrite     = $null
                importSucceeded = $false
                importExitCode  = $null
                importOutput    = $null
                verifyResults   = @()
            }

            if (-not (Test-Path -LiteralPath $regPath)) {
                $msg = "Strategy B: reg file missing for profile '$profName' at $regPath"
                $warnings += $msg
                Show-Warning "  $msg"
                $allProfilesVerified = $false
                $strategyBDetails += $perProfile
                break
            }

            # Phase 2.10.3 / extended in Phase 2.12.0: apply B-light pre-processing.
            # The 5 transforms inside cover: POP+IMAP binding strip (T2/T3),
            # user-path rewrite (T4), OST service-def drop (T5), and -- only
            # when version args differ -- cross-version path rewrite (T1).
            $blPath = Convert-RegFileToStrategyBLight `
                -SrcRegPath $regPath `
                -SourceUserName $blSrcUserName `
                -TargetUserName $blDstUserName `
                -SourceVersion $srcRegVer `
                -TargetVersion $tgtRegVer
            $perProfile.blTransformed = $true

            $sourcePrefix = Get-RegFileSourceHive -RegPath $blPath
            if ([string]::IsNullOrWhiteSpace($sourcePrefix)) {
                $msg = "Strategy B: could not detect source hive prefix in $regFile - falling back"
                $warnings += $msg
                Show-Warning "  $msg"
                if ($blPath -ne $regPath -and (Test-Path -LiteralPath $blPath)) {
                    Remove-Item -LiteralPath $blPath -Force -ErrorAction SilentlyContinue
                }
                $allProfilesVerified = $false
                $strategyBDetails += $perProfile
                break
            }
            $perProfile.hiveRewrite = if ($sourcePrefix -eq $targetHivePrefix) { 'none' }
                                     else { "$sourcePrefix -> $targetHivePrefix" }
            Show-Info "  [$profName] hive rewrite: $($perProfile.hiveRewrite)"

            $importPath = Convert-RegFileToTargetHive `
                -SrcRegPath $blPath `
                -SourcePrefix $sourcePrefix `
                -TargetPrefix $targetHivePrefix

            $importResult = Invoke-RegImport -RegPath $importPath
            $perProfile.importExitCode = $importResult.ExitCode
            $perProfile.importOutput   = $importResult.Output
            $perProfile.importSucceeded = $importResult.Success

            # Cleanup temp .reg files (both BL-transformed and hive-rewritten)
            foreach ($tmp in @($blPath, $importPath)) {
                if ($tmp -ne $regPath -and (Test-Path -LiteralPath $tmp)) {
                    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                }
            }

            if (-not $importResult.Success) {
                $msg = "Strategy B: reg.exe import failed for '$profName' (exit=$($importResult.ExitCode)): $($importResult.Output)"
                $warnings += $msg
                Show-Warning "  $msg"
                $allProfilesVerified = $false
                $strategyBDetails += $perProfile
                break
            }
            Show-Success "  [$profName] reg import OK"

            # Per-account verify
            $profileAccounts = @($plannedAccounts | Where-Object { $_.ProfileName -eq $profName })
            foreach ($pa in $profileAccounts) {
                $a = $pa.Account
                $vr = Test-AccountImported `
                    -HiveDrivePath $targetHivePsDrive `
                    -OutlookVersion $outlookVersion `
                    -ProfileName $profName `
                    -SubKey $a.subKey `
                    -ExpectedPop3Server $a.pop3.server
                $perProfile.verifyResults += [ordered]@{
                    subKey   = $a.subKey
                    email    = $a.email
                    verified = $vr.Verified
                    reason   = $vr.Reason
                }
                if ($vr.Verified) {
                    Show-Success "    [verify] $($a.email): OK"
                    # Reflect Strategy B success on the per-account result block
                    $accountResultRef = $resultsByAccount | Where-Object {
                        $_.profile -eq $profName -and $_.accountSubKey -eq $a.subKey
                    } | Select-Object -First 1
                    if ($null -ne $accountResultRef) {
                        $accountResultRef.verifyResult = 'verified'
                    }
                } else {
                    Show-Warning "    [verify] $($a.email): NG - $($vr.Reason)"
                    $warnings += "Strategy B verify failed for $profName/$($a.subKey) ($($a.email)): $($vr.Reason)"
                    $allProfilesVerified = $false
                }
            }

            $strategyBDetails += $perProfile
        }

        if ($allProfilesVerified) {
            $strategyBSucceeded = $true
            Show-Success 'Strategy B: all profiles imported and verified'
        } else {
            Show-Warning 'Strategy B: at least one profile/account failed - falling back to Strategy A'
        }
    }
}

# ----------------------------------------------------------
# Stage 5: operator communication
# ----------------------------------------------------------
$instructionsPath = $null

if ($strategyBSucceeded) {
    # ---- Stage 5a: Strategy B success popup ----
    $emails = ($resultsByAccount | Where-Object { $_.verifyResult -eq 'verified' } |
               ForEach-Object { $_.email }) -join ', '
    $popupTitle = 'Outlook POP - Restore Complete'
    $popupBody  = "$($plannedAccounts.Count) POP3 account(s) restored automatically:`r`n  $emails`r`n`r`n" +
                  "Operator action required (2 Outlook launches):`r`n" +
                  "  1. Launch Outlook. A 'restart required to link PST' notice will appear`r`n" +
                  "     and Outlook will close itself. This is expected behaviour - just close.`r`n" +
                  "  2. Launch Outlook again. Enter the password when prompted for each account.`r`n" +
                  "     Send/receive will work after password entry.`r`n" +
                  "     (DPAPI restriction: passwords cannot be migrated across machines)`r`n`r`n" +
                  "PST files and mail history are preserved. Contacts are visible from launch.`r`n`r`n" +
                  "Account settings (server / port / username / PST path) are saved`r`n" +
                  "as _account_settings.txt in the same folder as the PST file, in case`r`n" +
                  "manual re-setup is ever needed."
    if ($isCrossVersion) {
        $popupBody += "`r`n`r`n*** Cross-version restore ($crossVersionDirection) ***`r`n" +
                      "Additional manual cleanup steps are required on first launch:`r`n" +
                      "  - 'IMAP Search Folder' warning popup -> press OK`r`n" +
                      "  - Auto-created empty Outlook.pst -> set the migrated PST as default,`r`n" +
                      "    remove Outlook.pst, and use 'Change Folder' on the POP account`r`n" +
                      "  - Enter POP/IMAP passwords on first send/receive`r`n" +
                      "See _account_settings.txt 'Cross-version restore cleanup steps'`r`n" +
                      "section for the full step-by-step procedure."
    }
    try {
        Show-CompletionPopup -Title $popupTitle -Body $popupBody -Status 'Success'
    } catch { }
} else {
    # ---- Stage 5b: Strategy A fallback - RESTORE_INSTRUCTIONS.txt ----
    $instructionsPath = Join-Path $sectionDir 'RESTORE_INSTRUCTIONS.txt'
    try {
        $instructionsText = New-OutlookAccountInfoText `
            -Manifest $manifest `
            -TargetUserProfilePath $targetUserProfilePath `
            -PlannedAccounts $plannedAccounts `
            -ResultsByAccount $resultsByAccount `
            -StrategyBAttempted $strategyBAttempted `
            -StrategyBSucceeded $false `
            -IsCrossVersion $isCrossVersion `
            -CrossVersionDirection $crossVersionDirection
        $instructionsText | Out-File -FilePath $instructionsPath -Encoding UTF8 -Force
        Show-Info "Wrote instruction file: $instructionsPath"
    } catch {
        $warnings += "Failed to write instructions file: $($_.Exception.Message)"
        Show-Error "Failed to write instructions file: $($_.Exception.Message)"
    }

    try {
        $popupBody = "Manual setup is required for $($plannedAccounts.Count) Outlook POP account(s).`r`n`r`n" +
                     "PST file(s) have been placed at the target paths so that Outlook's wizard " +
                     "will attach them automatically when the operator adds the matching email account.`r`n`r`n" +
                     "Please open the instruction file and follow the steps:`r`n$instructionsPath"
        Show-CompletionPopup -Title 'Outlook POP - Manual Setup Required' -Body $popupBody -Status 'Partial'
    } catch { }

    try {
        Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$instructionsPath`"" -ErrorAction Stop | Out-Null
    } catch { }
}

# ----------------------------------------------------------
# Stage 5.5: always-on per-profile account-settings file in target
# Outlook Files folder. Travels with the PST so the operator (or a
# future engineer) can manually reconfigure the account if anything
# goes wrong. Written regardless of Strategy B success / Strategy A
# fallback. Filename: _account_settings.txt (leading underscore sorts
# above the PST file in Explorer alphabetical view).
# ----------------------------------------------------------
$targetSettingsWritten = @()
try {
    $profileDirMap = @{}
    foreach ($r in $resultsByAccount) {
        if ([string]::IsNullOrWhiteSpace($r.targetPstPath)) { continue }
        if ($r.status -ne 'Success') { continue }
        $dir = Split-Path -Path $r.targetPstPath -Parent
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        if (-not $profileDirMap.ContainsKey($r.profile)) {
            $profileDirMap[$r.profile] = $dir
        }
    }
    foreach ($profName in $profileDirMap.Keys) {
        $dir = $profileDirMap[$profName]
        if (-not (Test-Path -LiteralPath $dir)) {
            $warnings += "target dir missing for profile '$profName': $dir - skipping settings file"
            continue
        }
        $settingsPath = Join-Path $dir '_account_settings.txt'
        $settingsText = New-OutlookAccountInfoText `
            -Manifest $manifest `
            -TargetUserProfilePath $targetUserProfilePath `
            -PlannedAccounts $plannedAccounts `
            -ResultsByAccount $resultsByAccount `
            -StrategyBAttempted $strategyBAttempted `
            -StrategyBSucceeded $strategyBSucceeded `
            -ProfileFilter $profName `
            -IsCrossVersion $isCrossVersion `
            -CrossVersionDirection $crossVersionDirection
        $settingsText | Out-File -FilePath $settingsPath -Encoding UTF8 -Force
        $targetSettingsWritten += $settingsPath
        Show-Info "Wrote target settings file: $settingsPath"
    }
} catch {
    $warnings += "Failed to write target settings file(s): $($_.Exception.Message)"
    Show-Warning "Failed to write target settings file(s): $($_.Exception.Message)"
}

$sw.Stop()

$status = if ($strategyBSucceeded) {
    'Success'
} elseif ($failCount -gt 0 -and $successCount -eq 0) {
    'Failed'
} elseif ($failCount -gt 0) {
    'Partial'
} else {
    'Success'
}

return [PSCustomObject]@{
    Status               = $status
    ElapsedMs            = [int]$sw.ElapsedMilliseconds
    Summary              = [ordered]@{
        accountTotal           = $plannedAccounts.Count
        accountReady           = $successCount
        accountFail            = $failCount
        strategy               = if ($strategyBSucceeded) { 'B (reg-import)' }
                                 elseif ($strategyBAttempted) { 'A (fallback after B failure)' }
                                 else { 'A (no regExports in manifest)' }
        instructionsFile       = $instructionsPath
        targetSettingsFiles    = @($targetSettingsWritten)
        regProfilesImported    = @($strategyBDetails | Where-Object { $_.importSucceeded }).Count
        sourceOutlookFamily    = if ($sourceInstall) { $sourceInstall.productFamily } else { $null }
        targetOutlookFamily    = $targetInstall.ProductFamily
        targetOutlookType      = $targetInstall.InstallType
        targetOutlookExeVer    = $targetInstall.OutlookExeVersion
        crossVersionTransform  = [bool]$isCrossVersion
        crossVersionDirection  = $crossVersionDirection
    }
    Warnings             = $warnings
    ExternalOutputDir    = $null
    ExternalManifestPath = $null
    AccountResults       = @($resultsByAccount)
    StrategyBDetails     = @($strategyBDetails)
}
