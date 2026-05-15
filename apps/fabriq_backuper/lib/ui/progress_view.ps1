# ============================================================
# FabriqBackUper - Progress View
# Shows live (in Phase 2.1: completion-only) log + Done button.
# Note: Phase 2.1 runs ops synchronously on UI thread so the
# log only updates after completion. Phase 2.2 will move to a
# Runspace + timer for true live updates.
# ============================================================

$script:ProgressTitle    = $null
$script:ProgressLogBox   = $null
$script:ProgressDoneBtn  = $null

function New-ProgressView {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.BackColor = $script:bgForm

    $script:ProgressTitle = New-StyledLabel -Text "In progress..." `
        -X 24 -Y 14 -Width 800 -Height 28 -Font $script:fontLarge
    $panel.Controls.Add($script:ProgressTitle)

    # Log textbox (multiline, readonly, monospace) - Phase 2.7.1 compact.
    $log = New-Object System.Windows.Forms.TextBox
    $log.Multiline = $true
    $log.ReadOnly  = $true
    $log.ScrollBars = "Vertical"
    $log.Location = New-Object System.Drawing.Point(24, 50)
    $log.Size = New-Object System.Drawing.Size(880, 560)
    Set-TextBoxStyle -TextBox $log
    $log.Font = $script:fontMono
    $panel.Controls.Add($log)
    $script:ProgressLogBox = $log

    # Done button
    $btnDone = New-StyledButton -Text "Done" `
        -X 700 -Y 624 -Width 204 -Height 44 -BgColor $script:bgAccent
    $btnDone.Font = $script:fontLarge
    $btnDone.Enabled = $false
    $btnDone.Add_Click({ Switch-View 'ModeSelect' })
    $panel.Controls.Add($btnDone)
    $script:ProgressDoneBtn = $btnDone

    return $panel
}

function Initialize-ProgressView {
    param([string]$Title = "In progress...")
    if ($null -ne $script:ProgressTitle)   { $script:ProgressTitle.Text = $Title }
    if ($null -ne $script:ProgressLogBox)  { $script:ProgressLogBox.Text = "" }
    if ($null -ne $script:ProgressDoneBtn) { $script:ProgressDoneBtn.Enabled = $false }
}

function Add-ProgressLog {
    param([string]$Line)
    if ($null -eq $script:ProgressLogBox) { return }
    $script:ProgressLogBox.AppendText($Line + [Environment]::NewLine)
}

function Set-ProgressFinished {
    if ($null -ne $script:ProgressTitle)   { $script:ProgressTitle.Text = "Finished" }
    if ($null -ne $script:ProgressDoneBtn) { $script:ProgressDoneBtn.Enabled = $true }
}

# ============================================================
# Phase 2.7.4: summary formatting helpers used by Invoke-BackupStart
# and Invoke-RestoreStart to render the final-run summary block
# (elapsed time + aggregated data size).
# ============================================================

function Format-Bytes {
    param([Parameter(Mandatory = $true)][long]$Bytes)
    if ($Bytes -lt 0) { return "0 B" }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N2} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function Format-Duration {
    param([Parameter(Mandatory = $true)][TimeSpan]$Span)
    if ($Span.TotalHours -ge 1) {
        return ('{0}h {1:00}m {2:00}s' -f [int]$Span.TotalHours, $Span.Minutes, $Span.Seconds)
    }
    if ($Span.TotalMinutes -ge 1) {
        return ('{0}m {1:00}s' -f [int]$Span.TotalMinutes, $Span.Seconds)
    }
    return ('{0:N1}s' -f $Span.TotalSeconds)
}

# Phase 2.7.5: end-of-run completion popup. Shows a modal MessageBox so
# the operator notices the run finished even if they stepped away from
# the Progress View. Activate() pulls the form forward, MessageBox itself
# plays the OS notification sound + flashes the taskbar entry.
function Show-CompletionPopup {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Body,
        [Parameter(Mandatory = $true)][string]$Status  # Success / Partial / Failed / Skipped
    )
    $icon = switch ($Status) {
        'Failed'  { [System.Windows.Forms.MessageBoxIcon]::Error }
        'Partial' { [System.Windows.Forms.MessageBoxIcon]::Warning }
        default   { [System.Windows.Forms.MessageBoxIcon]::Information }
    }
    if ($null -ne $script:MainForm) { $script:MainForm.Activate() }
    [void][System.Windows.Forms.MessageBox]::Show(
        $script:MainForm,
        $Body,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon)
}

# Hook called by Switch-View when entering this view (no-op,
# caller is expected to initialize via Initialize-ProgressView).
function Show-ProgressView { }
