# ============================================================
# Easy Kitting Batch - Comment-Only Change Verifier
# ============================================================
# Verifies that a .ps1 file's modifications consist of comment
# changes ONLY (no functional code changes).
#
# Uses PowerShell's own parser to extract token streams, then
# compares everything except Comment-typed and NewLine-typed
# tokens. If the remaining token sequences are identical
# (Kind + Text), the change is provably comment-only and cannot
# affect runtime behavior.
#
# NewLine tokens are filtered alongside comments because:
#   1. They carry no PowerShell semantics distinct from `;` and
#      whitespace at the boundaries this verifier is designed to
#      check (translated comments may span more or fewer lines).
#   2. Cross-platform line-ending differences (LF vs CRLF) and
#      trailing-newline artifacts from `git show` extraction would
#      otherwise produce false FAILs on byte-identical content.
# A real code edit that, for example, replaces a NewLine with a
# Semi (`;`) token still changes the Semi token count and is
# detected.
#
# Usage:
#   # Mode 1: compare two arbitrary paths
#   powershell.exe -File ./dev/verify_comments_only.ps1 `
#       -Original kernel/common.ps1.bak `
#       -Modified kernel/common.ps1
#
#   # Mode 2: compare working tree against git HEAD
#   powershell.exe -File ./dev/verify_comments_only.ps1 -Path kernel/common.ps1
#
# Exit codes:
#   0 = PASS (only comment tokens differ)
#   1 = FAIL (non-comment tokens differ, parse error, or file not found)
#
# Notes:
# - Comment-Based Help dot-keywords (.SYNOPSIS / .PARAMETER / .EXAMPLE
#   etc.) are inside Comment tokens; the verifier does NOT detect
#   translation of those keywords. Translators must leave them intact
#   per project policy.
# - Mode 2 requires the file to be tracked in git HEAD.
# ============================================================

[CmdletBinding(DefaultParameterSetName = 'TwoPaths')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'TwoPaths')]
    [string]$Original,

    [Parameter(Mandatory = $true, ParameterSetName = 'TwoPaths')]
    [string]$Modified,

    [Parameter(Mandatory = $true, ParameterSetName = 'GitCompare')]
    [string]$Path
)

$ErrorActionPreference = 'Stop'

function Get-NonCommentTokens {
    param([string]$FilePath)
    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "File not found: $FilePath"
    }
    $absPath = (Resolve-Path -LiteralPath $FilePath).Path
    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $absPath, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $msg = ($errors | ForEach-Object {
            "L$($_.Extent.StartLineNumber): $($_.Message)"
        }) -join '; '
        throw "Parse errors in '${FilePath}': $msg"
    }
    return @($tokens | Where-Object {
        $_.Kind -ne 'Comment' -and $_.Kind -ne 'NewLine'
    })
}

function Compare-TokenStreams {
    param(
        [array]$OriginalTokens,
        [array]$ModifiedTokens,
        [string]$OriginalLabel,
        [string]$ModifiedLabel
    )

    if ($OriginalTokens.Count -ne $ModifiedTokens.Count) {
        Write-Host ("[FAIL] Significant token count differs: " +
                    "$OriginalLabel=$($OriginalTokens.Count), " +
                    "$ModifiedLabel=$($ModifiedTokens.Count)") `
                    -ForegroundColor Red
        return $false
    }

    for ($i = 0; $i -lt $OriginalTokens.Count; $i++) {
        $a = $OriginalTokens[$i]
        $b = $ModifiedTokens[$i]
        # Normalize CRLF -> LF before comparing Text. Some tokens
        # embed line terminators in their Text (LineContinuation,
        # multi-line StringExpandable / StringLiteral). Treating
        # CRLF and LF as equivalent makes the verifier robust to
        # cross-platform / cross-extraction line-ending differences
        # without sacrificing detection of real edits (a real edit
        # changes Kind, token sequence, or Text bytes beyond just
        # the `\r` byte).
        $aText = $a.Text -replace "`r`n", "`n"
        $bText = $b.Text -replace "`r`n", "`n"
        if ($a.Kind -ne $b.Kind -or $aText -ne $bText) {
            Write-Host "[FAIL] Token mismatch at index ${i}:" `
                -ForegroundColor Red
            Write-Host ("  $OriginalLabel L$($a.Extent.StartLineNumber): " +
                        "Kind=$($a.Kind), Text='$($a.Text)'") `
                        -ForegroundColor Red
            Write-Host ("  $ModifiedLabel L$($b.Extent.StartLineNumber): " +
                        "Kind=$($b.Kind), Text='$($b.Text)'") `
                        -ForegroundColor Red
            return $false
        }
    }
    return $true
}

function Invoke-GitShow {
    param(
        [string]$RelativePath,
        [string]$DestinationPath
    )
    # Read git's stdout as raw bytes (no encoding conversion) and
    # write them to the temp file verbatim. Earlier text-based
    # extraction (StandardOutput.ReadToEnd with UTF-8 decoder, then
    # WriteAllText with UTF-8 encoder) caused false FAILs on files
    # containing Japanese inside string literals: the round-trip
    # through .NET's UTF-16 string representation occasionally
    # shifted multi-byte boundaries and broke string termination
    # in the temp file's parse pass. BaseStream byte copy avoids
    # any decoder/encoder pairing entirely.
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'git'
    $startInfo.Arguments = "--no-pager show HEAD:$RelativePath"
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    # NOTE: do NOT set StandardOutputEncoding here. We access stdout
    # via BaseStream as bytes; setting the encoding would only matter
    # for text-based StandardOutput methods, which we are not using.
    $startInfo.StandardErrorEncoding = New-Object System.Text.UTF8Encoding $false
    $startInfo.WorkingDirectory = (Get-Location).Path

    $proc = [System.Diagnostics.Process]::Start($startInfo)

    # Copy stdout bytes verbatim into a memory stream
    $msOut = New-Object System.IO.MemoryStream
    try {
        $proc.StandardOutput.BaseStream.CopyTo($msOut)
    }
    finally {
        # Read stderr after stdout has drained to avoid pipe deadlock
        $stderr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()
    }

    if ($proc.ExitCode -ne 0) {
        throw ("git show HEAD:${RelativePath} failed (exit " +
                "$($proc.ExitCode)): $stderr")
    }

    [System.IO.File]::WriteAllBytes($DestinationPath, $msOut.ToArray())
}

try {
    if ($PSCmdlet.ParameterSetName -eq 'GitCompare') {
        if (-not (Test-Path -LiteralPath $Path)) {
            throw "File not found: $Path"
        }
        # Compute git-style relative path from repo root for `git show`
        $repoRoot = & git rev-parse --show-toplevel 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $repoRoot) {
            throw 'Not inside a git repository (git rev-parse failed)'
        }
        $absPath = (Resolve-Path -LiteralPath $Path).Path
        $relPath = $absPath.Substring($repoRoot.Length).TrimStart('\','/')
        $relPath = $relPath.Replace('\','/')

        $tmpFile = [System.IO.Path]::GetTempFileName()
        try {
            Invoke-GitShow -RelativePath $relPath -DestinationPath $tmpFile

            $origTokens = Get-NonCommentTokens -FilePath $tmpFile
            $modTokens  = Get-NonCommentTokens -FilePath $Path

            $ok = Compare-TokenStreams `
                -OriginalTokens $origTokens `
                -ModifiedTokens $modTokens `
                -OriginalLabel 'HEAD' `
                -ModifiedLabel 'WT'

            if (-not $ok) { exit 1 }
            Write-Host ("[PASS] ${Path}: $($origTokens.Count) " +
                        "significant tokens match HEAD exactly") `
                        -ForegroundColor Green
            exit 0
        }
        finally {
            if (Test-Path -LiteralPath $tmpFile) {
                Remove-Item -LiteralPath $tmpFile -Force
            }
        }
    }
    else {
        $origTokens = Get-NonCommentTokens -FilePath $Original
        $modTokens  = Get-NonCommentTokens -FilePath $Modified

        $ok = Compare-TokenStreams `
            -OriginalTokens $origTokens `
            -ModifiedTokens $modTokens `
            -OriginalLabel $Original `
            -ModifiedLabel $Modified

        if (-not $ok) { exit 1 }
        Write-Host ("[PASS] $Original vs ${Modified}: " +
                    "$($origTokens.Count) significant tokens match exactly") `
                    -ForegroundColor Green
        exit 0
    }
}
catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
