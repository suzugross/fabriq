# ========================================
# Manual Kitting Assistant
# ========================================
# Step-by-step manual kitting guidance GUI.
# Loads step_list.csv and prompt/*.txt to
# display work instructions for each step.
#
# NOTES:
# - Requires module to run inside fabriq
#   (kernel/common.ps1 must be dot-sourced)
# - phase 3 will add button actions:
#   Open, Copy1-3, Ctrl+C
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Ctrl+C send helper and no-activate button class for the floating UI
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class FabriqCtrlCSender {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    // keybd_event は private: C# ラッパー経由で呼び出す（PowerShell マーシャリング問題を回避）
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    public static void SendCtrlV() {
        const uint KEYEVENTF_KEYUP = 0x0002;
        keybd_event(0x11, 0, 0,               UIntPtr.Zero);  // Ctrl down
        keybd_event(0x56, 0, 0,               UIntPtr.Zero);  // V    down
        keybd_event(0x56, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);  // V    up
        keybd_event(0x11, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);  // Ctrl up
    }
}
public class NoActivateButton : Button {
    protected override void WndProc(ref Message m) {
        const int WM_MOUSEACTIVATE = 0x21;
        const int MA_NOACTIVATE    = 3;
        if (m.Msg == WM_MOUSEACTIVATE) {
            m.Result = new IntPtr(MA_NOACTIVATE);
            return;
        }
        base.WndProc(ref m);
    }
}
public class NoActivateForm : Form {
    // WS_EX_NOACTIVATE: クリックしてもウィンドウをアクティブにしない
    protected override CreateParams CreateParams {
        get {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
            return cp;
        }
    }
    // ベルト＋サスペンダー: WM_MOUSEACTIVATE でも MA_NOACTIVATE を返す
    protected override void WndProc(ref Message m) {
        const int WM_MOUSEACTIVATE = 0x21;
        const int MA_NOACTIVATE    = 3;
        if (m.Msg == WM_MOUSEACTIVATE) {
            m.Result = new IntPtr(MA_NOACTIVATE);
            return;
        }
        base.WndProc(ref m);
    }
}
'@ -ReferencedAssemblies "System.Windows.Forms" -ErrorAction SilentlyContinue

Write-Host ""
Show-Separator
Write-Host "Manual Kitting Assistant" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "step_list.csv"
$steps = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled","StepID","StepTitle","PromptFile",
                       "OpenCommand","OpenArgs","Copy1","Copy2","Copy3")

if ($null -eq $steps) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load step_list.csv")
}
if ($steps.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled steps")
}


# ========================================
# Step 2: Validate prompt/ directory and prompt files
# ========================================
$promptDir    = Join-Path $PSScriptRoot "prompt"
$hasAnyPrompt = $steps | Where-Object { -not [string]::IsNullOrWhiteSpace($_.PromptFile) }

if ($hasAnyPrompt -and (-not (Test-Path $promptDir))) {
    Show-Error "'prompt' directory not found: $promptDir"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "'prompt' directory not found")
}

$missingFiles = @()
foreach ($step in $steps) {
    if ([string]::IsNullOrWhiteSpace($step.PromptFile)) {
        continue  # Steps with empty PromptFile render in compact mode
    }
    $chkPath = Join-Path $promptDir $step.PromptFile
    if (-not (Test-Path $chkPath)) {
        $missingFiles += "$($step.StepID): $($step.PromptFile)"
    }
}
if ($missingFiles.Count -gt 0) {
    Show-Error "Missing prompt files:`n  $($missingFiles -join "`n  ")"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Missing prompt files: $($missingFiles.Count) file(s)")
}

Write-Host ""


# ========================================
# Step 3: Execution confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Start Manual Kitting Assistant? ($($steps.Count) steps)"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Gundam Light Theme - Color Palette
# ========================================
$bgForm     = [System.Drawing.Color]::FromArgb(240, 241, 245)  # mecha white
$bgPanel    = [System.Drawing.Color]::FromArgb(255, 255, 255)  # pure white panel
$bgPrompt   = [System.Drawing.Color]::FromArgb(248, 249, 252)  # text area background
$fgText     = [System.Drawing.Color]::FromArgb( 28,  32,  40)  # main text
$fgDim      = [System.Drawing.Color]::FromArgb(100, 110, 130)  # secondary text
$fgHeader   = [System.Drawing.Color]::FromArgb(255, 255, 255)  # header text
$borderClr  = [System.Drawing.Color]::FromArgb(200, 208, 220)  # border
$bgHeader   = [System.Drawing.Color]::FromArgb( 28,  43,  94)  # federation blue (deep navy)
$accentGold = [System.Drawing.Color]::FromArgb(200, 160,  30)  # V-fin gold
$btnRedBg   = [System.Drawing.Color]::FromArgb(190,  35,  28)  # gundam red (Execute)
$btnBlueBg  = [System.Drawing.Color]::FromArgb( 46,  95, 163)  # federal blue (Copy)
$btnYellowBg = [System.Drawing.Color]::FromArgb(170, 130,   0)  # gundam yellow (Ctrl+C)
$btnNavyBg  = [System.Drawing.Color]::FromArgb( 28,  43,  94)  # federation blue (Done)
$fgBtn      = [System.Drawing.Color]::FromArgb(255, 255, 255)  # button text


# ========================================
# Helper: Button factory
# ========================================
function New-StepButton {
    param(
        [string]$Text,
        [System.Drawing.Color]$BgColor,
        [int]$X = 10,
        [int]$Y = 0,
        [int]$W = 456,
        [int]$H = 38
    )
    $btn = New-Object NoActivateButton
    $btn.Text      = $Text
    $btn.Location  = New-Object System.Drawing.Point($X, $Y)
    $btn.Size      = New-Object System.Drawing.Size($W, $H)
    $btn.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize  = 0
    $btn.FlatAppearance.BorderColor = $BgColor
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
        [Math]::Min($BgColor.R + 28, 255),
        [Math]::Min($BgColor.G + 28, 255),
        [Math]::Min($BgColor.B + 28, 255)
    )
    $btn.BackColor = $BgColor
    $btn.ForeColor = $fgBtn
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    return $btn
}


# ========================================
# Form
# ========================================
$form = New-Object NoActivateForm
$form.Text            = "Manual Kitting Assistant"
$form.ClientSize      = New-Object System.Drawing.Size(476, 568)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
$form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual
$form.BackColor       = $bgForm
$form.ForeColor       = $fgText
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
$form.TopMost         = $true
$form.ShowInTaskbar   = $true
$form.KeyPreview      = $true

# Place at the bottom-right of the screen
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(
    ($screen.Right  - $form.Width  - 20),
    ($screen.Bottom - $form.Height - 20)
)


# ----------------------------------------
# Header panel (federation blue)
# ----------------------------------------
$pnlHeader           = New-Object System.Windows.Forms.Panel
$pnlHeader.Location  = New-Object System.Drawing.Point(0, 0)
$pnlHeader.Size      = New-Object System.Drawing.Size(476, 56)
$pnlHeader.BackColor = $bgHeader
$null = $form.Controls.Add($pnlHeader)

# Step progress (small text, pale blue)
$lblProgress          = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(12, 8)
$lblProgress.Size     = New-Object System.Drawing.Size(456, 18)
$lblProgress.Font     = New-Object System.Drawing.Font("Segoe UI", 8)
$lblProgress.ForeColor = [System.Drawing.Color]::FromArgb(170, 190, 225)
$lblProgress.BackColor = $bgHeader
$null = $pnlHeader.Controls.Add($lblProgress)

# Step title (bold white)
$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Location  = New-Object System.Drawing.Point(12, 28)
$lblTitle.Size      = New-Object System.Drawing.Size(456, 22)
$lblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $fgHeader
$lblTitle.BackColor = $bgHeader
$null = $pnlHeader.Controls.Add($lblTitle)


# ----------------------------------------
# V-fin gold accent line
# ----------------------------------------
$pnlAccent           = New-Object System.Windows.Forms.Panel
$pnlAccent.Location  = New-Object System.Drawing.Point(0, 56)
$pnlAccent.Size      = New-Object System.Drawing.Size(476, 3)
$pnlAccent.BackColor = $accentGold
$null = $form.Controls.Add($pnlAccent)


# ----------------------------------------
# Prompt display area
# ----------------------------------------
$pnlPromptBg           = New-Object System.Windows.Forms.Panel
$pnlPromptBg.Location  = New-Object System.Drawing.Point(0, 59)
$pnlPromptBg.Size      = New-Object System.Drawing.Size(476, 210)
$pnlPromptBg.BackColor = $bgPanel
$null = $form.Controls.Add($pnlPromptBg)

$txtPrompt              = New-Object System.Windows.Forms.TextBox
$txtPrompt.Location     = New-Object System.Drawing.Point(10, 10)
$txtPrompt.Size         = New-Object System.Drawing.Size(456, 190)
$txtPrompt.Font         = New-Object System.Drawing.Font("Segoe UI", 10)
$txtPrompt.ForeColor    = $fgText
$txtPrompt.BackColor    = $bgPrompt
$txtPrompt.ReadOnly     = $true
$txtPrompt.Multiline    = $true
$txtPrompt.WordWrap     = $true
$txtPrompt.BorderStyle  = [System.Windows.Forms.BorderStyle]::None
$txtPrompt.ScrollBars   = [System.Windows.Forms.ScrollBars]::Vertical
$txtPrompt.TabStop      = $false
$null = $pnlPromptBg.Controls.Add($txtPrompt)


# ----------------------------------------
# Separator line
# ----------------------------------------
$pnlSep           = New-Object System.Windows.Forms.Panel
$pnlSep.Location  = New-Object System.Drawing.Point(10, 273)
$pnlSep.Size      = New-Object System.Drawing.Size(456, 1)
$pnlSep.BackColor = $borderClr
$null = $form.Controls.Add($pnlSep)


# ----------------------------------------
# Action button group
# (initial Y is set dynamically by Update-StepDisplay)
# ----------------------------------------
$btnOpenCmd   = New-StepButton -Text "実行"           -BgColor $btnRedBg
$btnCopy1     = New-StepButton -Text "コピー1"        -BgColor $btnBlueBg
$btnCopy2     = New-StepButton -Text "コピー2"        -BgColor $btnBlueBg
$btnCopy3     = New-StepButton -Text "コピー3"        -BgColor $btnBlueBg
$btnCtrlCSend = New-Object NoActivateButton
$btnCtrlCSend.Text      = "貼り付け  (Ctrl+V)"
$btnCtrlCSend.Location  = New-Object System.Drawing.Point(10, 0)
$btnCtrlCSend.Size      = New-Object System.Drawing.Size(456, 38)
$btnCtrlCSend.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnCtrlCSend.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCtrlCSend.FlatAppearance.BorderSize  = 0
$btnCtrlCSend.FlatAppearance.BorderColor = $btnYellowBg
$btnCtrlCSend.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
    [Math]::Min($btnYellowBg.R + 28, 255),
    [Math]::Min($btnYellowBg.G + 28, 255),
    [Math]::Min($btnYellowBg.B + 28, 255)
)
$btnCtrlCSend.BackColor = $btnYellowBg
$btnCtrlCSend.ForeColor = $fgBtn
$btnCtrlCSend.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnCtrlCSend.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$btnCtrlCSend.Visible   = $true

$null = $form.Controls.Add($btnOpenCmd)
$null = $form.Controls.Add($btnCopy1)
$null = $form.Controls.Add($btnCopy2)
$null = $form.Controls.Add($btnCopy3)
$null = $form.Controls.Add($btnCtrlCSend)


# ----------------------------------------
# Done button panel (bottom, fixed position)
# ----------------------------------------
$pnlDone           = New-Object System.Windows.Forms.Panel
$pnlDone.Location  = New-Object System.Drawing.Point(0, 512)
$pnlDone.Size      = New-Object System.Drawing.Size(476, 56)
$pnlDone.BackColor = $bgPanel
$null = $form.Controls.Add($pnlDone)

# Thin separator above the Done button
$pnlDoneSep           = New-Object System.Windows.Forms.Panel
$pnlDoneSep.Location  = New-Object System.Drawing.Point(0, 0)
$pnlDoneSep.Size      = New-Object System.Drawing.Size(476, 1)
$pnlDoneSep.BackColor = $borderClr
$null = $pnlDone.Controls.Add($pnlDoneSep)

$btnDone = New-StepButton -Text "完了  (F2)" -BgColor $btnNavyBg -X 10 -Y 9 -W 456 -H 38
$null = $pnlDone.Controls.Add($btnDone)


# ========================================
# Update-StepDisplay
# Reflect a single CSV row's data across the entire UI
# ========================================
function Update-StepDisplay {
    param(
        [PSCustomObject]$Step,
        [int]$Index,
        [int]$Total
    )

    # -- Form title and header --
    $form.Text        = "Manual Kitting Assistant  [$Index / $Total]"
    $lblProgress.Text = "ステップ $($Step.StepID)   $Index / $Total"
    $lblTitle.Text    = $Step.StepTitle

    # -- Compact-mode decision --
    $hasPrompt           = -not [string]::IsNullOrWhiteSpace($Step.PromptFile)
    $pnlPromptBg.Visible = $hasPrompt
    $pnlSep.Visible      = $hasPrompt

    # -- Load prompt file --
    if ($hasPrompt) {
        $pPath = Join-Path $promptDir $Step.PromptFile
        try {
            $txtPrompt.Text = [System.IO.File]::ReadAllText($pPath, [System.Text.Encoding]::UTF8)
        }
        catch {
            $txtPrompt.Text = "(読み込みエラー: $($Step.PromptFile))"
            Show-Warning "Failed to read prompt file: $pPath"
        }
        $txtPrompt.SelectionStart = 0
        $txtPrompt.ScrollToCaret()
    }

    # -- Execute button (hidden when OpenCommand is empty) --
    $hasCmd = -not [string]::IsNullOrWhiteSpace($Step.OpenCommand)
    $btnOpenCmd.Visible = $hasCmd
    $btnOpenCmd.Tag = @{ Cmd = $Step.OpenCommand; Args = $Step.OpenArgs }

    # -- Copy buttons (each hidden when its value is empty) --
    @(
        @{ Btn = $btnCopy1; Val = $Step.Copy1 }
        @{ Btn = $btnCopy2; Val = $Step.Copy2 }
        @{ Btn = $btnCopy3; Val = $Step.Copy3 }
    ) | ForEach-Object {
        $hasVal = -not [string]::IsNullOrWhiteSpace($_.Val)
        $_.Btn.Visible = $hasVal
        $_.Btn.Tag = $_.Val
    }

    # -- Pack visible buttons from the top --
    $btnAreaTop = if ($hasPrompt) { 282 } else { 72 }
    $y          = $btnAreaTop
    $spacing    = 46   # button height (38px) + margin (8px)
    foreach ($btn in @($btnOpenCmd, $btnCopy1, $btnCopy2, $btnCopy3, $btnCtrlCSend)) {
        if ($btn.Visible) {
            $btn.Location = New-Object System.Drawing.Point(10, $y)
            $y += $spacing
        }
    }

    # -- Adjust form height dynamically --
    if ($hasPrompt) {
        # Normal mode: fixed height
        $form.ClientSize  = New-Object System.Drawing.Size(476, 568)
        $pnlDone.Location = New-Object System.Drawing.Point(0, 512)
    }
    else {
        # Compact mode: resize to fit the visible button count
        $pnlDoneY  = $y + 8
        $newHeight = $pnlDoneY + 56
        $form.ClientSize  = New-Object System.Drawing.Size(476, $newHeight)
        $pnlDone.Location = New-Object System.Drawing.Point(0, $pnlDoneY)
        # Reposition to the bottom-right of the screen
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $form.Location = New-Object System.Drawing.Point(
            ($screen.Right  - $form.Width  - 20),
            ($screen.Bottom - $form.Height - 20)
        )
    }
}


# ========================================
# Action state variable
# ========================================
$script:UserAction = $null   # "next" | "finished"


# ========================================
# Button event handlers
# ========================================

# Done button -> proceed to next step
$btnDone.Add_Click({ $script:UserAction = "next" })

# Keyboard shortcut: F2 = Done
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F2) {
        $script:UserAction = "next"
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
})

# X button -> cancel confirmation dialog
# gyotaq pattern: do not close the form here; let the main loop detect "cancel" and close it
$form.Add_FormClosing({
    param($sender, $e)
    if ($script:UserAction -ne "finished") {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "作業を中断してよろしいですか?`n`n残りのステップはスキップされます。",
            "中断の確認",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:UserAction = "cancel"
        }
        $e.Cancel = $true  # delegate the close to the main loop
    }
})

# -- Execute button --
# Tag = @{ Cmd = OpenCommand; Args = OpenArgs }
$btnOpenCmd.Add_Click({
    $tag = $btnOpenCmd.Tag
    if ($null -eq $tag -or [string]::IsNullOrWhiteSpace($tag.Cmd)) { return }
    try {
        if ([string]::IsNullOrWhiteSpace($tag.Args)) {
            $null = Start-Process -FilePath $tag.Cmd -PassThru
        } else {
            $null = Start-Process -FilePath $tag.Cmd -ArgumentList $tag.Args -PassThru
        }
        Start-Sleep -Milliseconds 500
    }
    catch {
        Show-Warning "Failed to open: $($tag.Cmd) - $_"
        [System.Windows.Forms.MessageBox]::Show(
            "起動に失敗しました:`n$($tag.Cmd)`n`n$_",
            "実行エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
})

# -- Copy buttons (shared handler) --
# Flash the button green on click as a visual cue that the copy succeeded
function Invoke-ClipboardCopy {
    param(
        [System.Windows.Forms.Button]$Btn,
        [System.Drawing.Color]$OriginalColor
    )
    $value = $Btn.Tag
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace($value)) { return }
    try {
        Set-Clipboard -Value $value
        $Btn.BackColor = [System.Drawing.Color]::FromArgb(34, 130, 60)  # green flash
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 600
        $Btn.BackColor = $OriginalColor
    }
    catch {
        Show-Warning "Failed to copy to clipboard: $_"
    }
}

$btnCopy1.Add_Click({ Invoke-ClipboardCopy -Btn $btnCopy1 -OriginalColor $btnBlueBg })
$btnCopy2.Add_Click({ Invoke-ClipboardCopy -Btn $btnCopy2 -OriginalColor $btnBlueBg })
$btnCopy3.Add_Click({ Invoke-ClipboardCopy -Btn $btnCopy3 -OriginalColor $btnBlueBg })

# -- Ctrl+V send --
# WS_EX_NOACTIVATE keeps the form from stealing focus on click.
# GetForegroundWindow returns the user-facing target window at click time
# (e.g., Edge). We re-assert focus to that target via SetForegroundWindow
# and then issue keybd_event through the C# wrapper SendCtrlV().
$btnCtrlCSend.Add_Click({
    $target = [FabriqCtrlCSender]::GetForegroundWindow()
    if ($target -eq [IntPtr]::Zero -or $target -eq $form.Handle) { return }
    try {
        [FabriqCtrlCSender]::SetForegroundWindow($target) | Out-Null
        [FabriqCtrlCSender]::SendCtrlV()
    }
    catch { Show-Warning "Failed to send Ctrl+C: $_" }
})


# ========================================
# Show the form (modeless)
# ========================================
$form.Show()


# ========================================
# Main loop
# ========================================
$totalSteps     = $steps.Count
$completedCount = 0
$skipCount      = 0
$current        = 0

foreach ($step in $steps) {
    $current++
    $script:UserAction = $null

    # Update the UI with the current step's data
    Update-StepDisplay -Step $step -Index $current -Total $totalSteps

    Show-Info "[$current/$totalSteps] $($step.StepID): $($step.StepTitle)"

    # Wait via DoEvents polling until the user presses Done or Cancel
    while ($null -eq $script:UserAction) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 50
    }

    if ($script:UserAction -eq "cancel") {
        # Mark the current step plus all remaining steps as Skip
        $skipCount += ($totalSteps - $current + 1)
        Show-Info "Cancelled by user. Remaining steps skipped."
        Write-Host ""
        break
    }

    $completedCount++
    Show-Success "Completed: $($step.StepTitle)"
    Write-Host ""
}


# ========================================
# Cleanup
# ========================================
$script:UserAction = "finished"
$form.Close()
$form.Dispose()


# ========================================
# Return result
# ========================================
return (New-BatchResult -Success $completedCount -Skip $skipCount -Fail 0 `
    -Title "Manual Kitting Assistant Results")
