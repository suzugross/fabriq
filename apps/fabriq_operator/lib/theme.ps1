# ========================================
# Fabriq Operator - Theme & UI Helpers
# ========================================
# CentreCOM-inspired light theme with a worn 2000s feel.
# Matches the fabriq_evidence_manager design language.
# ========================================

# ========================================
# Color Scheme (CentreCOM Light Theme)
# ========================================
$script:bgForm       = [System.Drawing.Color]::FromArgb(186, 190, 194)  # #BABEC2 window bg
$script:bgPanel      = [System.Drawing.Color]::FromArgb(74, 74, 74)     # #4A4A4A header bar
$script:bgGrid       = [System.Drawing.Color]::FromArgb(225, 228, 231)  # #E1E4E7 grid bg
$script:bgCellAlt    = [System.Drawing.Color]::FromArgb(237, 238, 239)  # #EDEEEF alt row
$script:bgCell       = [System.Drawing.Color]::FromArgb(225, 228, 231)  # #E1E4E7 normal row
$script:bgHeader     = [System.Drawing.Color]::FromArgb(160, 166, 171)  # #A0A6AB grid header
$script:bgButton     = [System.Drawing.Color]::FromArgb(152, 157, 161)  # #989DA1 button
$script:bgButtonHov  = [System.Drawing.Color]::FromArgb(170, 174, 179)  # #AAAEB3 hover
$script:bgAccent     = [System.Drawing.Color]::FromArgb(74, 144, 217)   # #4A90D9 accent blue
$script:bgAdd        = [System.Drawing.Color]::FromArgb(76, 175, 80)    # #4CAF50 success green
$script:bgDelete     = [System.Drawing.Color]::FromArgb(198, 40, 40)    # #C62828 error red
$script:bgInput      = [System.Drawing.Color]::FromArgb(255, 255, 255)  # white text input
$script:bgSelection  = [System.Drawing.Color]::FromArgb(76, 175, 80)    # #4CAF50 selected row
$script:bgTabPage    = [System.Drawing.Color]::FromArgb(196, 200, 204)  # #C4C8CC tab bg
$script:bgPreview    = [System.Drawing.Color]::FromArgb(210, 214, 219)  # #D2D6DB preview area

$script:fgText       = [System.Drawing.Color]::FromArgb(34, 34, 34)     # #222222 main text
$script:fgDim        = [System.Drawing.Color]::FromArgb(100, 100, 100)  # #646464 dim text
$script:fgHeader     = [System.Drawing.Color]::FromArgb(44, 44, 44)     # #2C2C2C header labels
$script:fgBtnText    = [System.Drawing.Color]::FromArgb(34, 34, 34)     # #222222 button text
$script:fgWhite      = [System.Drawing.Color]::FromArgb(255, 255, 255)  # white (on accent)
$script:fgGridHeader = [System.Drawing.Color]::FromArgb(44, 44, 44)     # #2C2C2C grid header text

$script:gridLine     = [System.Drawing.Color]::FromArgb(164, 168, 173)  # #A4A8AD grid lines
$script:borderColor  = [System.Drawing.Color]::FromArgb(117, 123, 130)  # #757B82 borders

# Header stripe accent colors
$script:stripeBlue   = [System.Drawing.Color]::FromArgb(74, 144, 217)   # #4A90D9
$script:stripeYellow = [System.Drawing.Color]::FromArgb(242, 201, 76)   # #F2C94C
$script:stripeRed    = [System.Drawing.Color]::FromArgb(235, 87, 87)    # #EB5757

# ========================================
# Fonts
# ========================================
$script:fontNormal   = New-UiFont "Segoe UI" 9
$script:fontBold     = New-UiFont "Segoe UI" 9 Bold
$script:fontSemiBold = New-UiFont "Segoe UI Semibold" 9
$script:fontLarge    = New-UiFont "Segoe UI" 11 Bold
$script:fontTitle    = New-UiFont "Segoe UI" 14 Bold
$script:fontMono     = New-UiFont "Consolas" 8.5

# ========================================
# Helper: Create a styled button
# ========================================
function New-StyledButton {
    param(
        [string]$Text,
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 120,
        [int]$Height = 30,
        $BgColor = $script:bgButton
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($Width, $Height)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor = $script:borderColor
    $btn.FlatAppearance.BorderSize = 1
    $btn.FlatAppearance.MouseOverBackColor = $script:bgButtonHov
    $btn.BackColor = $BgColor
    $btn.ForeColor = if ($BgColor -eq $script:bgAccent) { $script:fgWhite } else { $script:fgBtnText }
    $btn.Font = $script:fontNormal
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

# ========================================
# Helper: Style a DataGridView
# ========================================
function Set-GridStyle {
    param($Grid)
    $Grid.BackgroundColor = $script:bgGrid
    $Grid.GridColor = $script:gridLine
    $Grid.BorderStyle = "FixedSingle"
    $Grid.CellBorderStyle = "SingleHorizontal"
    $Grid.RowHeadersVisible = $false
    $Grid.AllowUserToAddRows = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.AllowUserToResizeRows = $false
    $Grid.ColumnHeadersHeightSizeMode = "DisableResizing"
    $Grid.ColumnHeadersHeight = 28
    $Grid.RowTemplate.Height = 26
    $Grid.DefaultCellStyle.BackColor = $script:bgCell
    $Grid.DefaultCellStyle.ForeColor = $script:fgText
    $Grid.DefaultCellStyle.SelectionBackColor = $script:bgSelection
    $Grid.DefaultCellStyle.SelectionForeColor = $script:fgWhite
    $Grid.DefaultCellStyle.Font = $script:fontNormal
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = $script:bgCellAlt
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:bgHeader
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:fgGridHeader
    $Grid.ColumnHeadersDefaultCellStyle.Font = $script:fontSemiBold
    $Grid.SelectionMode = "FullRowSelect"
    $Grid.MultiSelect = $false
    $Grid.ReadOnly = $true

    # DoubleBuffered for flicker-free rendering
    $t = $Grid.GetType()
    $p = $t.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
    $p.SetValue($Grid, $true, $null)
}

# ========================================
# Helper: Style a TextBox
# ========================================
function Set-TextBoxStyle {
    param($TextBox)
    $TextBox.BackColor = $script:bgInput
    $TextBox.ForeColor = $script:fgText
    $TextBox.Font = $script:fontNormal
    $TextBox.BorderStyle = "FixedSingle"
}

# ========================================
# Helper: Create a styled label
# ========================================
function New-StyledLabel {
    param(
        [string]$Text,
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 200,
        [int]$Height = 20,
        $FgColor = $script:fgText,
        $Font = $script:fontNormal
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    $lbl.Size = New-Object System.Drawing.Size($Width, $Height)
    $lbl.ForeColor = $FgColor
    $lbl.Font = $Font
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    return $lbl
}

# ========================================
# Helper: Create a styled panel
# ========================================
function New-StyledPanel {
    param(
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 100,
        [int]$Height = 100,
        $BgColor = $script:bgPanel
    )
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point($X, $Y)
    $panel.Size = New-Object System.Drawing.Size($Width, $Height)
    $panel.BackColor = $BgColor
    return $panel
}

# ========================================
# Helper: Style a CheckBox
# ========================================
function New-StyledCheckBox {
    param(
        [string]$Text,
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 160,
        [int]$Height = 22,
        [bool]$Checked = $false
    )
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $Text
    $cb.Location = New-Object System.Drawing.Point($X, $Y)
    $cb.Size = New-Object System.Drawing.Size($Width, $Height)
    $cb.ForeColor = $script:fgText
    $cb.Font = $script:fontNormal
    $cb.BackColor = [System.Drawing.Color]::Transparent
    $cb.Checked = $Checked
    return $cb
}

# ========================================
# Helper: Style a ComboBox
# ========================================
function New-StyledComboBox {
    param(
        [int]$X = 0,
        [int]$Y = 0,
        [int]$Width = 200,
        [int]$Height = 24
    )
    $cb = New-Object System.Windows.Forms.ComboBox
    $cb.Location = New-Object System.Drawing.Point($X, $Y)
    $cb.Size = New-Object System.Drawing.Size($Width, $Height)
    $cb.BackColor = $script:bgInput
    $cb.ForeColor = $script:fgText
    $cb.Font = $script:fontNormal
    $cb.DropDownStyle = "DropDownList"
    return $cb
}

# ========================================
# Helper: Apply standard form styling
# ========================================
function Set-FormStyle {
    param($Form, [string]$Title = "fabriq operator", [int]$Width = 700, [int]$Height = 520)
    $Form.Text = $Title
    # ClientSize, not Size: after the process turns DPI-aware (first
    # Capture-ScreenEvidence), window chrome gets taller/wider in pixels;
    # a fixed outer Size would let the grown chrome eat the client area
    # and clip the bottom row of controls. Pinning the CLIENT area keeps
    # every absolutely-positioned control visible in both DPI phases.
    # 16/39 = measured Win11 chrome at 96dpi (FixedSingle/FixedDialog/
    # Sizable alike), so unaware-phase geometry is unchanged.
    $Form.ClientSize = New-Object System.Drawing.Size(($Width - 16), ($Height - 39))
    $Form.StartPosition = "CenterScreen"
    $Form.BackColor = $script:bgForm
    $Form.ForeColor = $script:fgText
    $Form.Font = $script:fontNormal
    $Form.FormBorderStyle = "FixedSingle"
    $Form.MaximizeBox = $false
}
