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

# Hook called by Switch-View when entering this view (no-op,
# caller is expected to initialize via Initialize-ProgressView).
function Show-ProgressView { }
