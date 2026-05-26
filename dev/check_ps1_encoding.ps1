# ========================================
# .ps1 Encoding / Japanese-Content Checker
# ========================================
# Scans all .ps1 files under the project root and reports violations against
# the project conventions documented in:
#   memory/feedback_scripts_english_only.md
#     All .ps1 code (comments, string literals, identifiers) must be English only.
#   memory/feedback_ps1_utf8_bom.md
#     If Japanese is unavoidable, the .ps1 MUST be saved with UTF-8 BOM —
#     Windows PowerShell 5.1 reads BOM-less files as Shift_JIS on Japanese
#     Windows and mangles Japanese strings (was the cause of v1.0.0
#     printer_driver_install.ps1 regex-match failure, fixed in v1.1.0).
#
# Classification:
#   OK        : no Japanese, no BOM (compliant)
#   BOM-ONLY  : no Japanese, has BOM (cosmetic, not a violation)
#   JP+BOM    : Japanese present, has BOM (rule violation, encoding-safe)
#   JP-NO-BOM : Japanese present, no BOM (rule violation + encoding fragile)
#
# Exit code (Lenient mode, default):
#   0 — no JP-NO-BOM files
#   1 — one or more JP-NO-BOM files (encoding-fragile, must fix)
#
# Files in JP+BOM are reported as warnings but do NOT fail the run (legacy
# tolerance: existing apps like local_user_setup.ps1 are known violators per
# memory feedback_scripts_english_only).
#
# Usage:
#   pwsh ./dev/check_ps1_encoding.ps1
#   pwsh ./dev/check_ps1_encoding.ps1 -Root e:\fabriq
#
# Design note:
#   This script's source is intentionally ASCII-only. The Japanese/CJK Unicode
#   ranges are constructed via [char]0xXXXX casting rather than literal
#   characters, so the script does not fall into the very trap it detects.
# ========================================

param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = "Stop"

# Japanese / CJK Unicode ranges (composed via [char] casting, ASCII-only source):
#   U+3000-U+303F : CJK Symbols and Punctuation
#   U+3040-U+309F : Hiragana
#   U+30A0-U+30FF : Katakana
#   U+4E00-U+9FFF : CJK Unified Ideographs (kanji)
#   U+FF00-U+FFEF : Halfwidth and Fullwidth Forms
$jpPattern = '[' +
    [char]0x3000 + '-' + [char]0x303F +
    [char]0x3040 + '-' + [char]0x309F +
    [char]0x30A0 + '-' + [char]0x30FF +
    [char]0x4E00 + '-' + [char]0x9FFF +
    [char]0xFF00 + '-' + [char]0xFFEF +
    ']'

$rootFull = (Resolve-Path $Root).Path
if (-not (Test-Path $rootFull)) {
    Write-Host "[ERROR] Root path not found: $rootFull" -ForegroundColor Red
    exit 1
}

$ps1Files = Get-ChildItem -Path $rootFull -Recurse -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\\\.git\\' -and $_.FullName -notmatch '\\node_modules\\' }

$results = foreach ($f in $ps1Files) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)

    # Decode UTF-8 (BOM is the first character in the resulting string if present)
    $text = [System.Text.Encoding]::UTF8.GetString($bytes)
    if ($hasBom) { $text = $text.Substring(1) }

    $jpMatches = [regex]::Matches($text, $jpPattern)
    $jpCount = $jpMatches.Count

    $jpLines = @()
    if ($jpCount -gt 0) {
        $lineNum = 0
        foreach ($line in ($text -split "`r?`n")) {
            $lineNum++
            if ([regex]::IsMatch($line, $jpPattern)) {
                $jpLines += $lineNum
            }
        }
    }

    $rel = $f.FullName.Replace($rootFull + '\', '').Replace('\', '/')

    $level = if ($jpCount -eq 0 -and -not $hasBom) { 'OK' }
             elseif ($jpCount -eq 0 -and $hasBom) { 'BOM-ONLY' }
             elseif ($jpCount -gt 0 -and $hasBom) { 'JP+BOM' }
             else { 'JP-NO-BOM' }

    [PSCustomObject]@{
        Path        = $rel
        JpChars     = $jpCount
        HasBom      = $hasBom
        Level       = $level
        JpLineCount = $jpLines.Count
        FirstJpLine = if ($jpLines.Count -gt 0) { $jpLines[0] } else { $null }
    }
}

$total = $results.Count
$ok = @($results | Where-Object { $_.Level -eq 'OK' }).Count
$bomOnly = @($results | Where-Object { $_.Level -eq 'BOM-ONLY' }).Count
$jpWithBom = @($results | Where-Object { $_.Level -eq 'JP+BOM' }).Count
$jpNoBom = @($results | Where-Object { $_.Level -eq 'JP-NO-BOM' }).Count

Write-Host ""
Write-Host "========================================"
Write-Host " fabriq .ps1 encoding / JP-content check"
Write-Host "========================================"
Write-Host ""
Write-Host ("Root: {0}" -f $rootFull)
Write-Host ("Total .ps1 files: {0}" -f $total)
Write-Host ""
Write-Host ("  [OK]         no Japanese, no BOM (compliant)        : {0}" -f $ok)
Write-Host ("  [BOM-ONLY]   no Japanese, has BOM (cosmetic only)    : {0}" -f $bomOnly)
Write-Host ("  [JP+BOM]     Japanese + BOM (rule violation, WARN)   : {0}" -f $jpWithBom) -ForegroundColor Yellow
Write-Host ("  [JP-NO-BOM]  Japanese + no BOM (encoding fragile)    : {0}" -f $jpNoBom) -ForegroundColor Red
Write-Host ""

if ($jpNoBom -gt 0) {
    Write-Host "===== ERROR: JP-NO-BOM (encoding fragile, must fix) =====" -ForegroundColor Red
    $results | Where-Object { $_.Level -eq 'JP-NO-BOM' } |
        Sort-Object JpChars -Descending |
        Format-Table Path, JpChars, JpLineCount, FirstJpLine -AutoSize
}

if ($jpWithBom -gt 0) {
    Write-Host "===== WARNING: JP+BOM (rule violation, encoding-safe) =====" -ForegroundColor Yellow
    $results | Where-Object { $_.Level -eq 'JP+BOM' } |
        Sort-Object JpChars -Descending |
        Format-Table Path, JpChars, JpLineCount, FirstJpLine -AutoSize
}

if ($bomOnly -gt 0) {
    Write-Host "===== INFO: BOM-ONLY (no Japanese but has BOM) ====="
    $results | Where-Object { $_.Level -eq 'BOM-ONLY' } |
        Format-Table Path -AutoSize
}

if ($jpNoBom -gt 0) {
    Write-Host ""
    Write-Host ("[FAIL] {0} JP-NO-BOM violation(s) found (encoding fragile)" -f $jpNoBom) -ForegroundColor Red
    exit 1
}

Write-Host ""
if ($jpWithBom -gt 0) {
    Write-Host ("[OK with warnings] {0} JP+BOM violation(s) noted (legacy tolerance)" -f $jpWithBom) -ForegroundColor Yellow
} else {
    Write-Host "[OK] no Japanese content violations" -ForegroundColor Green
}
exit 0
