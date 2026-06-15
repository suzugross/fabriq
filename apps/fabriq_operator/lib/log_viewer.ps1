# ========================================
# Fabriq Operator - Module Log Viewer
# ========================================
# Presentation-only viewer over the existing per-module telemetry
# JSONL (logs/telemetry/<SessionID>/modules/<seq>_<name>.jsonl). It does
# NOT create a separate log stream; it reads what Show-* already records
# (type=show.<level>, tag, msg) and renders it color-coded so an operator
# can see how a module entry behaved without chasing the conhost window
# or the raw transcript (which misses async child-runspace output).
#
# Robustness contract (see t-0074 design):
#   - Reads ONLY the current $script:SessionID folder, so orphan folders
#     (created on resume) and other Profile runs cannot leak in.
#   - Identifies an entry by envelope.start.order (Profile Order), NOT the
#     filename sequence (which resets across __RESTART__ and would collide).
#   - When an Order ran multiple times this session, the newest run wins
#     (selected by file LastWriteTimeUtc, which is restart-proof).
#   - Malformed JSONL lines (partial writes) are skipped, never fatal.
#   - Telemetry schema is internal (dev/TELEMETRY_INTERNAL.md); parsing is
#     defensive (missing fields tolerated) to survive format drift.
# ========================================

# Maximum number of log lines rendered in the viewer. Very chatty modules
# (e.g. windows_update) can emit thousands of lines; beyond this cap the
# viewer shows a truncation notice instead of silently dropping them.
$script:LogViewerMaxLines = 5000

function Get-ModuleTelemetryLog {
    # Returns an array of [pscustomobject]@{ Level; Tag; Message; Ts } for
    # the newest telemetry run of the given Profile Order in the current
    # session. Returns an empty array when no telemetry was captured for
    # that Order (telemetry disabled, or the entry has not run yet).
    #
    # -ModulesDir is injectable for testing; in production it defaults to
    # the current session's telemetry modules folder.
    param(
        [Parameter(Mandatory)][int]$Order,
        [string]$ModulesDir = $null
    )

    if ([string]::IsNullOrWhiteSpace($ModulesDir)) {
        $sid = $script:SessionID
        if ([string]::IsNullOrWhiteSpace($sid)) { return ,@() }
        $ModulesDir = Join-Path (Join-Path ".\logs\telemetry" $sid) "modules"
    }

    if (-not (Test-Path -LiteralPath $ModulesDir)) { return ,@() }

    # Find every JSONL whose envelope.start.order matches. envelope.start is
    # the first event written per module, so the order is near the top; we
    # scan lines defensively rather than assuming a fixed position.
    $matchFiles = @()
    $files = @(Get-ChildItem -LiteralPath $ModulesDir -Filter "*.jsonl" -File -ErrorAction SilentlyContinue)
    foreach ($f in $files) {
        $lines = $null
        try { $lines = [System.IO.File]::ReadAllLines($f.FullName, [System.Text.Encoding]::UTF8) }
        catch { continue }
        $fileOrder = $null
        foreach ($ln in $lines) {
            if ([string]::IsNullOrWhiteSpace($ln)) { continue }
            $obj = $null
            try { $obj = $ln | ConvertFrom-Json } catch { continue }
            if ("$($obj.type)" -eq "envelope.start") {
                if ($null -ne $obj.order) {
                    try { $fileOrder = [int]$obj.order } catch { $fileOrder = $null }
                }
                break
            }
        }
        if ($null -ne $fileOrder -and $fileOrder -eq $Order) {
            $matchFiles += [pscustomobject]@{ File = $f; Lines = $lines }
        }
    }

    if ($matchFiles.Count -eq 0) { return ,@() }

    # Newest run wins. LastWriteTimeUtc is real wall-clock and therefore
    # survives __RESTART__ (where the in-process sequence counter resets).
    $latest = $matchFiles | Sort-Object { $_.File.LastWriteTimeUtc } -Descending | Select-Object -First 1

    $result = @()
    foreach ($ln in $latest.Lines) {
        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
        $obj = $null
        try { $obj = $ln | ConvertFrom-Json } catch { continue }
        $type = "$($obj.type)"
        if (-not $type.StartsWith("show.")) { continue }
        $result += [pscustomobject]@{
            Level   = $type.Substring(5)
            Tag     = "$($obj.tag)"
            Message = "$($obj.msg)"
            Ts      = "$($obj.ts)"
        }
    }

    # ,@() preserves an empty array at the caller's scalar assignment site
    # (a bare `return @()` collapses to $null). A populated array must NOT
    # be wrapped that way, or the unary comma nests it and the caller sees
    # a single element. So return the populated array plain.
    if ($result.Count -eq 0) { return ,@() }
    return $result
}

function Show-ModuleLogViewer {
    # Opens a modal Fabriq window showing the captured log for one module
    # entry. Non-destructive: it never mutates the FlexProfile dashboard
    # state; the caller stays on the dashboard after this returns.
    param(
        [Parameter(Mandatory)][int]$Order,
        [string]$ModuleName = ""
    )

    $lines = @(Get-ModuleTelemetryLog -Order $Order)

    $form = New-Object System.Windows.Forms.Form
    $titleName = if ([string]::IsNullOrWhiteSpace($ModuleName)) { "Order $Order" } else { "$ModuleName (Order $Order)" }
    Set-FormStyle -Form $form -Title "fabriq - Log: $titleName" -Width 820 -Height 580
    $form.FormBorderStyle = "Sizable"
    $form.MaximizeBox = $true

    # Header strip
    $headerPanel = New-StyledPanel -X 0 -Y 0 -Width 820 -Height 36 -BgColor $script:bgPanel
    $headerPanel.Anchor = "Top,Left,Right"
    $form.Controls.Add($headerPanel)
    $hdrLbl = New-StyledLabel -Text "Execution log: $titleName" -X 12 -Y 8 -Width 780 -Height 20 -Font $script:fontBold -FgColor $script:fgWhite
    $headerPanel.Controls.Add($hdrLbl)

    # Log body (read-only, monospaced, color-coded)
    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Location = New-Object System.Drawing.Point(10, 44)
    $rtb.Size = New-Object System.Drawing.Size(794, 462)
    $rtb.Anchor = "Top,Bottom,Left,Right"
    $rtb.ReadOnly = $true
    $rtb.BackColor = $script:bgInput
    $rtb.Font = $script:fontMono
    $rtb.WordWrap = $false
    $rtb.ScrollBars = "Both"
    $rtb.BorderStyle = "FixedSingle"
    $form.Controls.Add($rtb)

    # Level -> text color. Chosen for readability on the white body bg
    # (the theme's bgAdd/stripeYellow are tuned for fills, not text).
    $colError   = [System.Drawing.Color]::FromArgb(198, 40, 40)
    $colWarning = [System.Drawing.Color]::FromArgb(180, 95, 6)
    $colSuccess = [System.Drawing.Color]::FromArgb(46, 125, 50)
    $colSkip    = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $colInfo    = $script:fgText

    $appendLine = {
        param($text, $color)
        $rtb.SelectionStart = $rtb.TextLength
        $rtb.SelectionLength = 0
        $rtb.SelectionColor = $color
        $rtb.AppendText($text + "`n")
    }

    if ($lines.Count -eq 0) {
        & $appendLine "No log captured for this entry." $colSkip
        & $appendLine "(Telemetry may be disabled, or this entry has not run in the current session.)" $colSkip
    }
    else {
        $shown = $lines
        $truncated = $false
        if ($lines.Count -gt $script:LogViewerMaxLines) {
            $shown = $lines[0..($script:LogViewerMaxLines - 1)]
            $truncated = $true
        }
        foreach ($l in $shown) {
            $color = switch ("$($l.Level)") {
                'error'   { $colError }
                'warning' { $colWarning }
                'success' { $colSuccess }
                'skip'    { $colSkip }
                default   { $colInfo }
            }
            $tag = "$($l.Level)".ToUpper()
            & $appendLine ("[{0}] {1}" -f $tag, $l.Message) $color
        }
        if ($truncated) {
            $dropped = $lines.Count - $script:LogViewerMaxLines
            & $appendLine "" $colInfo
            & $appendLine ("... truncated: $dropped more line(s). See the raw telemetry JSONL for the full log.") $colWarning
        }
    }
    $rtb.SelectionStart = 0
    $rtb.ScrollToCaret()

    # Footer: redaction note + Close
    $noteLbl = New-StyledLabel -Text "Values are redacted; see the raw transcript for unmasked detail." -X 12 -Y 516 -Width 560 -Height 20 -FgColor $script:fgDim
    $noteLbl.Anchor = "Bottom,Left"
    $form.Controls.Add($noteLbl)

    $btnClose = New-StyledButton -Text "Close" -X 694 -Y 512 -Width 110 -Height 28
    $btnClose.Anchor = "Bottom,Right"
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    $form.Dispose()
}
