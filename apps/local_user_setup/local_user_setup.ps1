# ========================================
# Local User Setup GUI for Fabriq
# ========================================
# local_user_list.csv のプレースホルダー行に
# ユーザー名とパスワードを順番に登録するウィザード型 GUI。
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Color Scheme (Light / Bright)
# ========================================
$bgForm       = [System.Drawing.Color]::FromArgb(245, 247, 250)
$bgSection    = [System.Drawing.Color]::FromArgb(255, 255, 255)
$bgInput      = [System.Drawing.Color]::FromArgb(255, 255, 255)
$borderColor  = [System.Drawing.Color]::FromArgb(210, 218, 230)
$fgText       = [System.Drawing.Color]::FromArgb(30,  40,  50)
$fgDim        = [System.Drawing.Color]::FromArgb(120, 130, 145)
$fgHeader     = [System.Drawing.Color]::FromArgb(0,   90,  170)
$fgValue      = [System.Drawing.Color]::FromArgb(20,  30,  40)
$fgError      = [System.Drawing.Color]::FromArgb(196, 43,  28)
$fgSuccess    = [System.Drawing.Color]::FromArgb(0,   130, 60)
$fgWarning    = [System.Drawing.Color]::FromArgb(180, 90,  0)
$bgPrimary    = [System.Drawing.Color]::FromArgb(0,   120, 215)
$bgPrimaryDis = [System.Drawing.Color]::FromArgb(190, 205, 225)
$fgPrimaryDis = [System.Drawing.Color]::FromArgb(130, 145, 165)
$fgButton     = [System.Drawing.Color]::White

# ========================================
# CSV Path (fixed)
# ========================================
$script:csvPath = Join-Path $PSScriptRoot "..\..\modules\standard\local_user_config\local_user_list.csv"
$script:csvPath = [System.IO.Path]::GetFullPath($script:csvPath)

# ========================================
# State
# ========================================
$script:allRows      = @()
$script:placeholders = @()
$script:currentIdx   = 0

# ========================================
# Helper: Label
# ========================================
function New-Label {
    param(
        [string]$Text,
        [int]$X, [int]$Y, [int]$W, [int]$H,
        $Color = $fgText,
        $Font  = $null,
        [string]$Align = "MiddleLeft"
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text
    $lbl.Location  = New-Object System.Drawing.Point($X, $Y)
    $lbl.Size      = New-Object System.Drawing.Size($W, $H)
    $lbl.ForeColor = $Color
    $lbl.TextAlign = $Align
    if ($Font) { $lbl.Font = $Font }
    return $lbl
}

# ========================================
# Helper: Separator line
# ========================================
function New-Separator {
    param([int]$Y)
    $sep = New-Object System.Windows.Forms.Panel
    $sep.Location  = New-Object System.Drawing.Point(20, $Y)
    $sep.Size      = New-Object System.Drawing.Size(490, 1)
    $sep.BackColor = $borderColor
    return $sep
}

# ========================================
# Helper: TextBox
# ========================================
function New-InputBox {
    param([int]$X, [int]$Y, [int]$W = 370, [bool]$Password = $false)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location    = New-Object System.Drawing.Point($X, $Y)
    $txt.Size        = New-Object System.Drawing.Size($W, 28)
    $txt.BackColor   = $bgInput
    $txt.ForeColor   = $fgText
    $txt.BorderStyle = "FixedSingle"
    $txt.Font        = New-Object System.Drawing.Font("Segoe UI", 10)
    if ($Password) { $txt.PasswordChar = '*' }
    return $txt
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Local User Setup - Fabriq"
$form.Size            = New-Object System.Drawing.Size(540, 530)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $bgForm
$form.ForeColor       = $fgText
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false

# ========================================
# SECTION 1: Slot Info
# ========================================

# --- Header + Progress ---
$lblSlotHeader = New-Label -Text "次に作成するユーザー枠" `
    -X 20 -Y 18 -W 320 -H 24 -Color $fgHeader `
    -Font (New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold))
$form.Controls.Add($lblSlotHeader)

$lblProgress = New-Label -Text "" -X 360 -Y 18 -W 150 -H 24 -Color $fgHeader `
    -Font (New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)) `
    -Align "MiddleRight"
$form.Controls.Add($lblProgress)

$form.Controls.Add((New-Separator -Y 47))

# --- Slot info panel (white card) ---
$pnlSlot = New-Object System.Windows.Forms.Panel
$pnlSlot.Location  = New-Object System.Drawing.Point(20, 55)
$pnlSlot.Size      = New-Object System.Drawing.Size(490, 120)
$pnlSlot.BackColor = $bgSection
$form.Controls.Add($pnlSlot)

# "No slot" message (hidden initially)
$lblNoSlot = New-Label -Text "作成可能なユーザー枠がありません。`nlocal_user_list.csv にプレースホルダー行を追加してください。" `
    -X 15 -Y 22 -W 460 -H 70 -Color $fgWarning `
    -Font (New-Object System.Drawing.Font("Segoe UI", 10))
$lblNoSlot.Visible = $false
$pnlSlot.Controls.Add($lblNoSlot)

# Slot info rows (caption + value)
$slotRows = @(
    @{ Key = "Group";      Caption = "グループ";       Y = 10 }
    @{ Key = "PwdExp";     Caption = "パスワード期限"; Y = 36 }
    @{ Key = "PwdChange";  Caption = "パスワード変更"; Y = 62 }
    @{ Key = "Desc";       Caption = "説明";           Y = 88 }
)
$slotLabels = @{}

foreach ($row in $slotRows) {
    $cap = New-Label -Text "$($row.Caption):" -X 16 -Y $row.Y -W 120 -H 24 -Color $fgDim
    $pnlSlot.Controls.Add($cap)

    $val = New-Label -Text "" -X 140 -Y $row.Y -W 340 -H 24 -Color $fgValue `
        -Font (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold))
    $pnlSlot.Controls.Add($val)

    $slotLabels[$row.Key] = $val
}

$form.Controls.Add((New-Separator -Y 183))

# ========================================
# SECTION 2: Input
# ========================================

$lblInputHeader = New-Label -Text "ユーザー情報を入力" `
    -X 20 -Y 197 -W 300 -H 24 -Color $fgHeader `
    -Font (New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold))
$form.Controls.Add($lblInputHeader)

# --- UserName ---
$lblUserName = New-Label -Text "ユーザー名:" -X 20 -Y 235 -W 110 -H 28 -Color $fgText
$form.Controls.Add($lblUserName)
$txtUserName = New-InputBox -X 135 -Y 233 -W 375
$form.Controls.Add($txtUserName)

# --- Password ---
$lblPassword = New-Label -Text "パスワード:" -X 20 -Y 275 -W 110 -H 28 -Color $fgText
$form.Controls.Add($lblPassword)
$txtPassword = New-InputBox -X 135 -Y 273 -W 375 -Password $true
$form.Controls.Add($txtPassword)

# --- Confirm Password ---
$lblConfirm = New-Label -Text "確認入力:" -X 20 -Y 315 -W 110 -H 28 -Color $fgText
$form.Controls.Add($lblConfirm)
$txtConfirm = New-InputBox -X 135 -Y 313 -W 375 -Password $true
$form.Controls.Add($txtConfirm)

# --- Show password checkbox ---
$chkShow = New-Object System.Windows.Forms.CheckBox
$chkShow.Text      = "パスワードを表示する"
$chkShow.Location  = New-Object System.Drawing.Point(135, 350)
$chkShow.Size      = New-Object System.Drawing.Size(200, 22)
$chkShow.ForeColor = $fgDim
$chkShow.BackColor = $bgForm
$chkShow.FlatStyle = "Flat"
$form.Controls.Add($chkShow)

$form.Controls.Add((New-Separator -Y 384))

# ========================================
# SECTION 3: Action + Status
# ========================================

# --- Register button ---
$btnRegister = New-Object System.Windows.Forms.Button
$btnRegister.Text      = "登録して次へ  ▶"
$btnRegister.Location  = New-Object System.Drawing.Point(135, 398)
$btnRegister.Size      = New-Object System.Drawing.Size(220, 40)
$btnRegister.FlatStyle = "Flat"
$btnRegister.FlatAppearance.BorderSize = 0
$btnRegister.BackColor = $bgPrimary
$btnRegister.ForeColor = $fgButton
$btnRegister.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnRegister.Cursor    = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnRegister)

# --- Close button (shown only when all registrations are complete) ---
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text      = "完了  ✓"
$btnClose.Location  = New-Object System.Drawing.Point(135, 398)
$btnClose.Size      = New-Object System.Drawing.Size(220, 40)
$btnClose.FlatStyle = "Flat"
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 80)
$btnClose.ForeColor = $fgButton
$btnClose.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnClose.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnClose.Visible   = $false
$form.Controls.Add($btnClose)

# --- Status label ---
$lblStatus = New-Label -Text "準備完了" -X 20 -Y 454 -W 490 -H 22 -Color $fgDim -Align "MiddleLeft"
$form.Controls.Add($lblStatus)

# ========================================
# Logic: Placeholder extraction
# ========================================
function Get-Placeholders {
    param($Rows)
    return @($Rows | Where-Object {
        $_.Enabled -eq "1" -and
        [string]::IsNullOrWhiteSpace($_.UserName) -and
        [string]::IsNullOrWhiteSpace($_.Password) -and
        -not [string]::IsNullOrWhiteSpace($_.Group)
    })
}

# ========================================
# Logic: Load current placeholder into labels
# ========================================
function Load-Placeholder {
    $ph    = $script:placeholders[$script:currentIdx]
    $total = $script:placeholders.Count

    $lblProgress.Text = "$($script:currentIdx + 1) 枠目  /  全 $total 枠"

    $slotLabels["Group"].Text     = $ph.Group
    $slotLabels["PwdExp"].Text    = if ($ph.PasswordNeverExpires -eq "1")      { "永続（期限なし）" } else { "期限あり" }
    $slotLabels["PwdChange"].Text = if ($ph.UserMayNotChangePassword -eq "1")  { "変更禁止"         } else { "変更可"   }
    $slotLabels["Desc"].Text      = if ([string]::IsNullOrWhiteSpace($ph.Description)) { "（なし）" } else { $ph.Description }

    $lblNoSlot.Visible = $false
}

# ========================================
# Logic: Show exhausted / error state
# ========================================
function Set-ExhaustedState {
    param(
        [string]$Message = "作成可能なユーザー枠がありません。`nlocal_user_list.csv にプレースホルダー行を追加してください。",
        [switch]$ShowClose
    )

    $lblProgress.Text = ""
    foreach ($key in $slotLabels.Keys) { $slotLabels[$key].Text = "" }
    $lblNoSlot.Text    = $Message
    $lblNoSlot.Visible = $true

    $txtUserName.Enabled  = $false
    $txtPassword.Enabled  = $false
    $txtConfirm.Enabled   = $false
    $chkShow.Enabled      = $false
    $btnRegister.Enabled  = $false
    $btnRegister.BackColor = $bgPrimaryDis
    $btnRegister.ForeColor = $fgPrimaryDis

    $lblStatus.Text      = "登録できる枠がありません"
    $lblStatus.ForeColor = $fgWarning

    if ($ShowClose) {
        $btnRegister.Visible = $false
        $btnClose.Visible    = $true
    }
}

# ========================================
# Logic: Clear inputs and focus
# ========================================
function Clear-Inputs {
    $txtUserName.Text = ""
    $txtPassword.Text = ""
    $txtConfirm.Text  = ""
    $txtUserName.Focus() | Out-Null
    $lblStatus.Text      = "準備完了"
    $lblStatus.ForeColor = $fgDim
}

# ========================================
# Events
# ========================================

# --- Form Load ---
$form.Add_Load({
    # CSV existence check
    if (-not (Test-Path $script:csvPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "CSV ファイルが見つかりません:`n$($script:csvPath)",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        Set-ExhaustedState -Message "CSV ファイルが見つかりません。`n$($script:csvPath)"
        return
    }

    # Load CSV
    try {
        $script:allRows      = @(Import-Csv -Path $script:csvPath -Encoding UTF8)
        $script:placeholders = Get-Placeholders -Rows $script:allRows
        $script:currentIdx   = 0
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "CSV の読み込みに失敗しました:`n$($_.Exception.Message)",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        Set-ExhaustedState -Message "CSV の読み込みに失敗しました。"
        return
    }

    if ($script:placeholders.Count -eq 0) {
        Set-ExhaustedState
    }
    else {
        Load-Placeholder
        Clear-Inputs
    }
})

# --- Show/Hide password ---
$chkShow.Add_CheckedChanged({
    $char = if ($chkShow.Checked) { [char]0 } else { '*' }
    $txtPassword.PasswordChar = $char
    $txtConfirm.PasswordChar  = $char
})

# --- Enter key navigation ---
$txtUserName.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $txtPassword.Focus() | Out-Null
    }
})

$txtPassword.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $txtConfirm.Focus() | Out-Null
    }
})

$txtConfirm.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $btnRegister.PerformClick()
    }
})

# --- Register button ---
$btnRegister.Add_Click({

    # ---- Validation ----
    if ([string]::IsNullOrWhiteSpace($txtUserName.Text)) {
        $lblStatus.Text      = "ユーザー名を入力してください"
        $lblStatus.ForeColor = $fgError
        $txtUserName.Focus() | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($txtPassword.Text)) {
        $lblStatus.Text      = "パスワードを入力してください"
        $lblStatus.ForeColor = $fgError
        $txtPassword.Focus() | Out-Null
        return
    }

    if ($txtPassword.Text -ne $txtConfirm.Text) {
        $lblStatus.Text      = "パスワードが一致しません — 確認入力をもう一度入力してください"
        $lblStatus.ForeColor = $fgError
        $txtConfirm.Text     = ""
        $txtConfirm.Focus() | Out-Null
        return
    }

    # ---- Apply to CSV row (in-memory) ----
    $script:placeholders[$script:currentIdx].UserName = $txtUserName.Text.Trim()
    $script:placeholders[$script:currentIdx].Password = $txtPassword.Text

    # ---- Save CSV ----
    try {
        $script:allRows | Export-Csv -Path $script:csvPath -NoTypeInformation -Encoding UTF8 -Force
    }
    catch {
        # Revert in-memory change to keep state consistent
        $script:placeholders[$script:currentIdx].UserName = ""
        $script:placeholders[$script:currentIdx].Password = ""
        [System.Windows.Forms.MessageBox]::Show(
            "CSV の保存に失敗しました:`n$($_.Exception.Message)",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    # ---- Success feedback ----
    $registeredName          = $txtUserName.Text.Trim()
    $lblStatus.Text          = "登録完了:  $registeredName"
    $lblStatus.ForeColor     = $fgSuccess

    # ---- Advance to next placeholder ----
    $script:currentIdx++

    if ($script:currentIdx -ge $script:placeholders.Count) {
        Set-ExhaustedState -Message "すべてのユーザー枠への登録が完了しました。" -ShowClose
        $lblStatus.Text      = "すべての枠への登録が完了しました"
        $lblStatus.ForeColor = $fgSuccess
    }
    else {
        Load-Placeholder
        Clear-Inputs
    }
})

# --- Close button ---
$btnClose.Add_Click({
    $form.Close()
})

# ========================================
# Show Form
# ========================================
$form.ShowDialog() | Out-Null
$form.Dispose()
