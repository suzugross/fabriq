# ========================================
# System Launcher
# ========================================
# Quick access to Windows settings, control panel, and system tools
# without using Windows Search (no search history left behind).
# Provides both GUI (with search) and console menu modes.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:tools = @(
    # --- Settings (ms-settings) ---
    [PSCustomObject]@{ Num=1;  Name="Display";            Command="ms-settings:display";          Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=2;  Name="Sound";              Command="ms-settings:sound";             Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=3;  Name="Notifications";      Command="ms-settings:notifications";     Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=4;  Name="Network";            Command="ms-settings:network";           Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=5;  Name="Wi-Fi";              Command="ms-settings:network-wifi";      Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=6;  Name="Proxy";              Command="ms-settings:network-proxy";     Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=7;  Name="Apps & Features";    Command="ms-settings:appsfeatures";      Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=8;  Name="Default Apps";       Command="ms-settings:defaultapps";       Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=9;  Name="Taskbar";            Command="ms-settings:taskbar";           Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=10; Name="Personalization";    Command="ms-settings:personalization";   Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=11; Name="Date & Time";        Command="ms-settings:dateandtime";       Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=12; Name="Region";             Command="ms-settings:regionformatting";  Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=13; Name="Windows Update";     Command="ms-settings:windowsupdate";     Type="uri";     Category="Settings" }
    [PSCustomObject]@{ Num=14; Name="About";              Command="ms-settings:about";             Type="uri";     Category="Settings" }

    # --- Control Panel ---
    [PSCustomObject]@{ Num=15; Name="Control Panel (All)"; Command="control.exe";                  Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=16; Name="Network Connections"; Command="ncpa.cpl";                     Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=17; Name="Programs & Features"; Command="appwiz.cpl";                   Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=18; Name="System Properties";   Command="sysdm.cpl";                    Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=19; Name="Power Options";       Command="powercfg.cpl";                 Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=20; Name="Sound (Legacy)";      Command="mmsys.cpl";                    Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=21; Name="Firewall";            Command="firewall.cpl";                 Type="exe";     Category="Control Panel" }
    [PSCustomObject]@{ Num=22; Name="User Accounts";       Command="netplwiz";                     Type="exe";     Category="Control Panel" }

    # --- System Tools ---
    [PSCustomObject]@{ Num=23; Name="God Mode";            Command="shell:::{ED7BA470-8E54-465E-825C-99712043E01C}"; Type="shell"; Category="System Tools" }
    [PSCustomObject]@{ Num=24; Name="Device Manager";      Command="devmgmt.msc";                  Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=25; Name="Computer Management"; Command="compmgmt.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=26; Name="Event Viewer";        Command="eventvwr.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=27; Name="Services";            Command="services.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=28; Name="Task Scheduler";      Command="taskschd.msc";                 Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=29; Name="Group Policy";        Command="gpedit.msc";                   Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=30; Name="Registry Editor";     Command="regedit.exe";                  Type="exe";     Category="System Tools" }
    [PSCustomObject]@{ Num=31; Name="Disk Management";     Command="diskmgmt.msc";                 Type="exe";     Category="System Tools" }

    # --- Shell ---
    [PSCustomObject]@{ Num=32; Name="CMD";                 Command="cmd.exe";                      Type="exe";     Category="Shell" }
    [PSCustomObject]@{ Num=33; Name="PowerShell";          Command="powershell.exe";               Type="exe";     Category="Shell" }
    [PSCustomObject]@{ Num=34; Name="PowerShell (Admin)";  Command="powershell.exe";               Type="runas";   Category="Shell" }
)

function Invoke-Tool {
    param([PSCustomObject]$Tool)
    switch ($Tool.Type) {
        "uri"   { Start-Process $Tool.Command -ErrorAction SilentlyContinue }
        "shell" { Start-Process explorer.exe $Tool.Command -ErrorAction SilentlyContinue }
        "runas" { Start-Process $Tool.Command -Verb RunAs -ErrorAction SilentlyContinue }
        "exe"   { Start-Process $Tool.Command -ErrorAction SilentlyContinue }
    }
}

# ========================================
# GUI Launcher
# ========================================

$bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
$bgCell       = [System.Drawing.Color]::FromArgb(45, 45, 45)
$bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
$bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
$bgLaunch     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
$fgPlaceholder = [System.Drawing.Color]::FromArgb(100, 100, 100)
$gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)

function Update-ToolGrid {
    param([string]$Filter = "")
    $script:dgv.Rows.Clear()

    foreach ($t in $script:tools) {
        if (-not [string]::IsNullOrWhiteSpace($Filter)) {
            $match = ($t.Name -like "*$Filter*") -or
                     ($t.Category -like "*$Filter*") -or
                     ($t.Command -like "*$Filter*")
            if (-not $match) { continue }
        }

        $idx = $script:dgv.Rows.Add()
        $row = $script:dgv.Rows[$idx]
        $row.Cells["Num"].Value      = $t.Num
        $row.Cells["ToolName"].Value = $t.Name
        $row.Cells["Category"].Value = $t.Category
        $row.Cells["Command"].Value  = $t.Command
        $row.Tag = $t
    }

    $countText = "$($script:dgv.Rows.Count) / $($script:tools.Count) items"
    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        $countText += " (filtered)"
    }
    $script:statusLabel.Text = $countText
}

function Invoke-SelectedTool {
    if ($script:dgv.CurrentRow -and $script:dgv.CurrentRow.Tag) {
        $tool = $script:dgv.CurrentRow.Tag
        $script:statusLabel.Text = "Launching: $($tool.Name)..."
        try {
            Invoke-Tool -Tool $tool
            $script:statusLabel.Text = "Launched: $($tool.Name)"
        }
        catch {
            $script:statusLabel.Text = "Failed: $($tool.Name)"
        }
    }
}

# ----- Main Form -----
$form = New-Object System.Windows.Forms.Form
$form.Text = "Fabriq System Launcher"
$form.Size = New-Object System.Drawing.Size(750, 550)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgDark
$form.ForeColor = $fgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

# ----- Top Panel (Search + Launch) -----
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Top"
$topPanel.Height = 50
$topPanel.BackColor = $bgPanel
$topPanel.Padding = New-Object System.Windows.Forms.Padding(10)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search:"
$searchLabel.Location = New-Object System.Drawing.Point(12, 15)
$searchLabel.AutoSize = $true
$searchLabel.ForeColor = $fgDim

$script:searchBox = New-Object System.Windows.Forms.TextBox
$script:searchBox.Location = New-Object System.Drawing.Point(70, 12)
$script:searchBox.Size = New-Object System.Drawing.Size(300, 26)
$script:searchBox.BackColor = $bgCell
$script:searchBox.ForeColor = $fgText
$script:searchBox.BorderStyle = "FixedSingle"
$script:searchBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$script:searchBox.Add_TextChanged({
    Update-ToolGrid -Filter $script:searchBox.Text
})

$btnLaunch = New-Object System.Windows.Forms.Button
$btnLaunch.Text = "Launch"
$btnLaunch.Size = New-Object System.Drawing.Size(90, 30)
$btnLaunch.FlatStyle = "Flat"
$btnLaunch.FlatAppearance.BorderColor = $gridLine
$btnLaunch.FlatAppearance.MouseOverBackColor = $bgButtonHov
$btnLaunch.BackColor = $bgLaunch
$btnLaunch.ForeColor = $fgText
$btnLaunch.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnLaunch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnLaunch.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 108), 10)

$form.Add_Resize({
    $btnLaunch.Left = $form.ClientSize.Width - 108
})

$btnLaunch.Add_Click({ Invoke-SelectedTool })

$topPanel.Controls.AddRange(@($searchLabel, $script:searchBox, $btnLaunch))
$form.Controls.Add($topPanel)

# ----- Status Bar -----
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = "Bottom"
$statusPanel.Height = 28
$statusPanel.BackColor = $bgPanel

$script:statusLabel = New-Object System.Windows.Forms.Label
$script:statusLabel.Text = "Ready"
$script:statusLabel.Location = New-Object System.Drawing.Point(10, 7)
$script:statusLabel.AutoSize = $true
$script:statusLabel.ForeColor = $fgDim

$statusPanel.Controls.Add($script:statusLabel)
$form.Controls.Add($statusPanel)

# ----- DataGridView -----
$script:dgv = New-Object System.Windows.Forms.DataGridView
$script:dgv.Dock = "Fill"
$script:dgv.BackgroundColor = $bgGrid
$script:dgv.GridColor = $gridLine
$script:dgv.BorderStyle = "None"
$script:dgv.CellBorderStyle = "SingleHorizontal"
$script:dgv.RowHeadersVisible = $false
$script:dgv.AllowUserToAddRows = $false
$script:dgv.AllowUserToDeleteRows = $false
$script:dgv.AllowUserToResizeRows = $false
$script:dgv.SelectionMode = "FullRowSelect"
$script:dgv.MultiSelect = $false
$script:dgv.ReadOnly = $true
$script:dgv.AutoSizeColumnsMode = "Fill"
$script:dgv.ColumnHeadersHeightSizeMode = "DisableResizing"
$script:dgv.ColumnHeadersHeight = 32
$script:dgv.RowTemplate.Height = 28
$script:dgv.DefaultCellStyle.BackColor = $bgCell
$script:dgv.DefaultCellStyle.ForeColor = $fgText
$script:dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 80, 140)
$script:dgv.DefaultCellStyle.SelectionForeColor = $fgText
$script:dgv.EnableHeadersVisualStyles = $false
$script:dgv.ColumnHeadersDefaultCellStyle.BackColor = $bgHeader
$script:dgv.ColumnHeadersDefaultCellStyle.ForeColor = $fgHeader
$script:dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$dgvType = $script:dgv.GetType()
$pi = $dgvType.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
$pi.SetValue($script:dgv, $true, $null)

# Columns
$colNum = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNum.Name = "Num"
$colNum.HeaderText = "#"
$colNum.FillWeight = 8
$null = $script:dgv.Columns.Add($colNum)

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = "ToolName"
$colName.HeaderText = "Name"
$colName.FillWeight = 30
$null = $script:dgv.Columns.Add($colName)

$colCat = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCat.Name = "Category"
$colCat.HeaderText = "Category"
$colCat.FillWeight = 20
$null = $script:dgv.Columns.Add($colCat)

$colCmd = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCmd.Name = "Command"
$colCmd.HeaderText = "Command"
$colCmd.FillWeight = 42
$null = $script:dgv.Columns.Add($colCmd)

$form.Controls.Add($script:dgv)

# ----- Events -----
$script:dgv.Add_CellDoubleClick({
    param($s, $e)
    if ($e.RowIndex -ge 0) { Invoke-SelectedTool }
})

$script:dgv.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.Handled = $true
        Invoke-SelectedTool
    }
})

$script:searchBox.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.Handled = $true
        $e.SuppressKeyPress = $true
        if ($script:dgv.Rows.Count -gt 0) {
            $script:dgv.CurrentCell = $script:dgv.Rows[0].Cells["ToolName"]
            Invoke-SelectedTool
        }
    }
    elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Down) {
        $script:dgv.Focus()
        if ($script:dgv.Rows.Count -gt 0) {
            $script:dgv.CurrentCell = $script:dgv.Rows[0].Cells["ToolName"]
        }
    }
})

# ----- Layout & Show -----
$statusPanel.BringToFront()
$topPanel.BringToFront()
$script:dgv.BringToFront()

Update-ToolGrid

Show-Info "Opening System Launcher GUI..."
Write-Host ""

$form.ShowDialog() | Out-Null
$form.Dispose()

Show-Info "System Launcher closed"
Write-Host ""
