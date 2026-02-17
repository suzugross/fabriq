# ========================================
# Store App Editor for Fabriq
# ========================================
# GUI tool for browsing installed Store apps
# and editing storeapp_list.csv for the
# Store App Removal module.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ========================================
# Color Scheme (Fabriq Standard)
# ========================================
$bgDark       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel      = [System.Drawing.Color]::FromArgb(40, 40, 40)
$bgGrid       = [System.Drawing.Color]::FromArgb(35, 35, 35)
$bgCell       = [System.Drawing.Color]::FromArgb(45, 45, 45)
$bgHeader     = [System.Drawing.Color]::FromArgb(55, 55, 55)
$bgButton     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgButtonHov  = [System.Drawing.Color]::FromArgb(80, 80, 80)
$bgAccent     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$bgAdd        = [System.Drawing.Color]::FromArgb(40, 140, 60)
$bgDelete     = [System.Drawing.Color]::FromArgb(180, 40, 40)
$fgText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
$fgDim        = [System.Drawing.Color]::FromArgb(150, 150, 150)
$fgHeader     = [System.Drawing.Color]::FromArgb(100, 180, 255)
$gridLine     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgInput      = [System.Drawing.Color]::FromArgb(50, 50, 50)

# ========================================
# Default CSV Path
# ========================================
$script:defaultCsvPath = Join-Path $PSScriptRoot "..\..\modules\standard\storeapp_config\storeapp_list.csv"
$script:defaultCsvPath = [System.IO.Path]::GetFullPath($script:defaultCsvPath)

# ========================================
# State Tracking
# ========================================
$script:isDirty = $false
$script:currentCsvPath = $null
$script:allInstalledApps = @()

# ========================================
# Helper: Styled Button
# ========================================
function New-StyledButton {
    param([string]$Text, [int]$X, [int]$Y = 0, [int]$Width = 100, [int]$Height = 28, $BgColor = $bgButton)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($Width, $Height)
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderColor = $gridLine
    $btn.FlatAppearance.MouseOverBackColor = $bgButtonHov
    $btn.BackColor = $BgColor
    $btn.ForeColor = $fgText
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

# ========================================
# Helper: Style a DataGridView
# ========================================
function Set-GridStyle {
    param($Grid)
    $Grid.BackgroundColor = $bgGrid
    $Grid.GridColor = $gridLine
    $Grid.BorderStyle = "None"
    $Grid.CellBorderStyle = "SingleHorizontal"
    $Grid.RowHeadersVisible = $false
    $Grid.AllowUserToAddRows = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.AllowUserToResizeRows = $false
    $Grid.ColumnHeadersHeightSizeMode = "DisableResizing"
    $Grid.ColumnHeadersHeight = 30
    $Grid.RowTemplate.Height = 26
    $Grid.DefaultCellStyle.BackColor = $bgCell
    $Grid.DefaultCellStyle.ForeColor = $fgText
    $Grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 80, 140)
    $Grid.DefaultCellStyle.SelectionForeColor = $fgText
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $bgHeader
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = $fgHeader
    $Grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    # DoubleBuffered
    $t = $Grid.GetType()
    $p = $t.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
    $p.SetValue($Grid, $true, $null)
}

# ========================================
# Helper: Style a TextBox
# ========================================
function Set-TextBoxStyle {
    param($TextBox)
    $TextBox.BackColor = $bgInput
    $TextBox.ForeColor = $fgText
    $TextBox.BorderStyle = "FixedSingle"
}

# ========================================
# Helper: Load CSV into DataTable
# ========================================
function Load-CsvToTable {
    param([string]$Path, [System.Data.DataTable]$Table)

    $csvData = Import-Csv -Path $Path -Encoding Default
    $Table.Rows.Clear()

    foreach ($item in $csvData) {
        $props = $item.PSObject.Properties.Name

        $no      = ""
        $appName = ""
        $enabled = ""
        $desc    = ""

        if ("No" -in $props)          { $no = $item.No }
        if ("AppName" -in $props)     { $appName = $item.AppName }
        if ("Enabled" -in $props)     { $enabled = $item.Enabled }
        if ("Description" -in $props) { $desc = $item.Description }

        if ([string]::IsNullOrWhiteSpace($enabled)) { $enabled = "1" }

        $Table.Rows.Add($no, $appName, $enabled, $desc)
    }
}

# ========================================
# Helper: Filter installed apps grid
# ========================================
function Update-InstalledFilter {
    param([string]$FilterText)

    $gridInstalled.Rows.Clear()
    $filter = $FilterText.Trim().ToLower()

    foreach ($app in $script:allInstalledApps) {
        if ($filter -and -not ($app.Name.ToLower().Contains($filter))) {
            continue
        }
        $null = $gridInstalled.Rows.Add($app.Name, $app.PackageFullName)
    }

    $lblInstalledCount.Text = "Showing: $($gridInstalled.Rows.Count) / $($script:allInstalledApps.Count)"
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Store App Editor - Fabriq"
$form.Size = New-Object System.Drawing.Size(1050, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgDark
$form.ForeColor = $fgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ========================================
# SplitContainer (Top: Installed, Bottom: CSV)
# ========================================
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = "Fill"
$splitContainer.Orientation = "Horizontal"
$splitContainer.SplitterDistance = 380
$splitContainer.SplitterWidth = 5
$splitContainer.BackColor = $bgDark
$splitContainer.Panel1.BackColor = $bgDark
$splitContainer.Panel2.BackColor = $bgDark

# ==========================================
# TOP PANEL: Installed Apps
# ==========================================
$panelTop = $splitContainer.Panel1
$panelTop.Padding = New-Object System.Windows.Forms.Padding(10)

# --- Header Label ---
$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Text = "Installed Store Apps"
$lblHeader.Location = New-Object System.Drawing.Point(10, 8)
$lblHeader.Size = New-Object System.Drawing.Size(160, 20)
$lblHeader.ForeColor = $fgHeader
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelTop.Controls.Add($lblHeader)

# --- Filter Row ---
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Location = New-Object System.Drawing.Point(10, 36)
$lblFilter.Size = New-Object System.Drawing.Size(45, 20)
$lblFilter.ForeColor = $fgText
$panelTop.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(58, 34)
$txtFilter.Size = New-Object System.Drawing.Size(350, 25)
Set-TextBoxStyle $txtFilter
$panelTop.Controls.Add($txtFilter)

$btnRefresh = New-StyledButton -Text "Refresh" -X 415 -Y 32 -Width 90 -BgColor $bgAccent
$panelTop.Controls.Add($btnRefresh)

$btnAdd = New-StyledButton -Text "Add to CSV >>" -X 515 -Y 32 -Width 120 -BgColor $bgAdd
$panelTop.Controls.Add($btnAdd)

# --- Installed count label ---
$lblInstalledCount = New-Object System.Windows.Forms.Label
$lblInstalledCount.Location = New-Object System.Drawing.Point(650, 36)
$lblInstalledCount.Size = New-Object System.Drawing.Size(350, 18)
$lblInstalledCount.Text = ""
$lblInstalledCount.ForeColor = $fgDim
$lblInstalledCount.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$panelTop.Controls.Add($lblInstalledCount)

# --- Installed Apps Grid ---
$gridInstalled = New-Object System.Windows.Forms.DataGridView
$gridInstalled.Location = New-Object System.Drawing.Point(10, 65)
$gridInstalled.Size = New-Object System.Drawing.Size(1010, 250)
$gridInstalled.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$gridInstalled.SelectionMode = "FullRowSelect"
$gridInstalled.MultiSelect = $true
$gridInstalled.ReadOnly = $true
$gridInstalled.AutoSizeColumnsMode = "Fill"
Set-GridStyle $gridInstalled

$colIName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colIName.Name = "Name"
$colIName.HeaderText = "Name (AppName)"
$colIName.FillWeight = 40
$null = $gridInstalled.Columns.Add($colIName)

$colIFullName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colIFullName.Name = "PackageFullName"
$colIFullName.HeaderText = "PackageFullName"
$colIFullName.FillWeight = 60
$null = $gridInstalled.Columns.Add($colIFullName)

$panelTop.Controls.Add($gridInstalled)

# --- Description + Add Row (below grid) ---
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text = "Description:"
$lblDesc.Location = New-Object System.Drawing.Point(10, 325)
$lblDesc.Size = New-Object System.Drawing.Size(80, 20)
$lblDesc.ForeColor = $fgText
$panelTop.Controls.Add($lblDesc)

$txtDesc = New-Object System.Windows.Forms.TextBox
$txtDesc.Location = New-Object System.Drawing.Point(93, 323)
$txtDesc.Size = New-Object System.Drawing.Size(350, 25)
Set-TextBoxStyle $txtDesc
$panelTop.Controls.Add($txtDesc)

$lblDescHint = New-Object System.Windows.Forms.Label
$lblDescHint.Text = "(empty = use AppName)"
$lblDescHint.Location = New-Object System.Drawing.Point(450, 325)
$lblDescHint.Size = New-Object System.Drawing.Size(200, 18)
$lblDescHint.ForeColor = $fgDim
$panelTop.Controls.Add($lblDescHint)

# ==========================================
# BOTTOM PANEL: CSV Editor
# ==========================================
$panelBottom = $splitContainer.Panel2
$panelBottom.Padding = New-Object System.Windows.Forms.Padding(10)

# --- CSV Toolbar ---
$lblCsv = New-Object System.Windows.Forms.Label
$lblCsv.Text = "CSV Editor"
$lblCsv.Location = New-Object System.Drawing.Point(10, 10)
$lblCsv.Size = New-Object System.Drawing.Size(80, 20)
$lblCsv.ForeColor = $fgHeader
$lblCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelBottom.Controls.Add($lblCsv)

$btnLoad      = New-StyledButton -Text "Load CSV"   -X 100  -Y 6 -Width 90 -BgColor $bgAccent
$btnSave      = New-StyledButton -Text "Save CSV"   -X 200  -Y 6 -Width 90 -BgColor $bgAccent
$btnDeleteRow = New-StyledButton -Text "Delete Row"  -X 300  -Y 6 -Width 100 -BgColor $bgDelete

$panelBottom.Controls.Add($btnLoad)
$panelBottom.Controls.Add($btnSave)
$panelBottom.Controls.Add($btnDeleteRow)

# --- CSV file path label ---
$lblCsvPath = New-Object System.Windows.Forms.Label
$lblCsvPath.Location = New-Object System.Drawing.Point(410, 10)
$lblCsvPath.Size = New-Object System.Drawing.Size(580, 18)
$lblCsvPath.Text = "(No file loaded)"
$lblCsvPath.ForeColor = $fgDim
$lblCsvPath.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$panelBottom.Controls.Add($lblCsvPath)

# --- CSV DataTable ---
$dt = New-Object System.Data.DataTable
[void]$dt.Columns.Add("No")
[void]$dt.Columns.Add("AppName")
[void]$dt.Columns.Add("Enabled")
[void]$dt.Columns.Add("Description")

# --- CSV Grid ---
$gridCsv = New-Object System.Windows.Forms.DataGridView
$gridCsv.Location = New-Object System.Drawing.Point(10, 38)
$gridCsv.Size = New-Object System.Drawing.Size(1010, 250)
$gridCsv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$gridCsv.AutoSizeColumnsMode = "Fill"
$gridCsv.SelectionMode = "FullRowSelect"
$gridCsv.MultiSelect = $true
Set-GridStyle $gridCsv
$gridCsv.AllowUserToAddRows = $true
$gridCsv.AutoGenerateColumns = $false
$gridCsv.DataSource = $dt

# No column
$colNo = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNo.HeaderText = "No"
$colNo.DataPropertyName = "No"
$colNo.Name = "No"
$colNo.FillWeight = 1
$colNo.MinimumWidth = 30
$null = $gridCsv.Columns.Add($colNo)

# AppName column
$colAppName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colAppName.HeaderText = "AppName"
$colAppName.DataPropertyName = "AppName"
$colAppName.Name = "AppName"
$colAppName.FillWeight = 6
$colAppName.MinimumWidth = 80
$null = $gridCsv.Columns.Add($colAppName)

# Enabled column as ComboBox (0/1 dropdown)
$colEnabled = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colEnabled.HeaderText = "Enabled"
$colEnabled.DataPropertyName = "Enabled"
$colEnabled.Name = "Enabled"
$colEnabled.Items.AddRange("1", "0")
$colEnabled.FlatStyle = "Flat"
$colEnabled.FillWeight = 1
$colEnabled.MinimumWidth = 50
$null = $gridCsv.Columns.Add($colEnabled)

# Description column
$colDesc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDesc.HeaderText = "Description"
$colDesc.DataPropertyName = "Description"
$colDesc.Name = "Description"
$colDesc.FillWeight = 30
$null = $gridCsv.Columns.Add($colDesc)

$panelBottom.Controls.Add($gridCsv)

# ========================================
# Assemble Layout
# ========================================
$form.Controls.Add($splitContainer)

# ========================================
# Helper: Load installed apps
# ========================================
function Load-InstalledApps {
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblInstalledCount.Text = "Loading..."

    try {
        $script:allInstalledApps = @(Get-AppxPackage | Select-Object Name, PackageFullName | Sort-Object Name)
        Update-InstalledFilter -FilterText $txtFilter.Text
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load installed apps: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $lblInstalledCount.Text = "Load failed"
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

# ========================================
# Events
# ========================================

# --- Form Load: Installed apps + Default CSV ---
$form.Add_Load({
    Load-InstalledApps

    # Auto-load default CSV if exists
    if (Test-Path $script:defaultCsvPath) {
        try {
            Load-CsvToTable -Path $script:defaultCsvPath -Table $dt
            $script:currentCsvPath = $script:defaultCsvPath
            $lblCsvPath.Text = $script:defaultCsvPath
            $script:isDirty = $false
        }
        catch {
            $lblCsvPath.Text = "(Failed to load default CSV)"
        }
    }
})

# --- Filter TextBox: Real-time filtering ---
$txtFilter.Add_TextChanged({
    Update-InstalledFilter -FilterText $txtFilter.Text
})

# --- Refresh Button ---
$btnRefresh.Add_Click({
    Load-InstalledApps
})

# --- Add to CSV Button ---
$btnAdd.Add_Click({
    if ($gridInstalled.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select app(s) from the installed list first.",
            "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $addedCount = 0
    $skippedCount = 0

    # Determine next No value
    $maxNo = 0
    foreach ($row in $dt.Rows) {
        $n = 0
        if ([int]::TryParse($row["No"], [ref]$n)) {
            if ($n -gt $maxNo) { $maxNo = $n }
        }
    }

    foreach ($selRow in $gridInstalled.SelectedRows) {
        $appName = $selRow.Cells["Name"].Value

        if ([string]::IsNullOrWhiteSpace($appName)) { continue }

        # Duplicate check
        $existing = $dt.Select("AppName = '$($appName -replace "'","''")'")
        if ($existing.Count -gt 0) {
            $skippedCount++
            continue
        }

        $maxNo += 10
        $desc = $txtDesc.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($desc)) { $desc = $appName }
        $dt.Rows.Add($maxNo.ToString(), $appName, "1", $desc)
        $addedCount++
    }

    if ($addedCount -gt 0) {
        $script:isDirty = $true
    }

    if ($skippedCount -gt 0) {
        $lblInstalledCount.Text = "Added: $addedCount, Skipped (duplicate): $skippedCount"
    }
    elseif ($addedCount -gt 0) {
        $lblInstalledCount.Text = "Added: $addedCount"
    }
})

# --- Delete Row Button ---
$btnDeleteRow.Add_Click({
    if ($gridCsv.SelectedRows.Count -eq 0) { return }

    $indices = @()
    foreach ($row in $gridCsv.SelectedRows) {
        if ($row.Index -lt $dt.Rows.Count) {
            $indices += $row.Index
        }
    }
    $indices = $indices | Sort-Object -Descending

    foreach ($idx in $indices) {
        $dt.Rows.RemoveAt($idx)
    }
    $script:isDirty = $true
})

# --- Load CSV Button ---
$btnLoad.Add_Click({
    if ($script:isDirty) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Unsaved changes will be lost. Continue?",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $ofd.Title = "Load CSV"

    if ($script:currentCsvPath) {
        $ofd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:currentCsvPath)
    }
    elseif (Test-Path $script:defaultCsvPath) {
        $ofd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:defaultCsvPath)
    }

    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        Load-CsvToTable -Path $ofd.FileName -Table $dt
        $script:currentCsvPath = $ofd.FileName
        $lblCsvPath.Text = $ofd.FileName
        $script:isDirty = $false
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load CSV: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

# --- Save CSV Button ---
$btnSave.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $sfd.Title = "Save CSV"
    $sfd.FileName = "storeapp_list.csv"

    if ($script:currentCsvPath) {
        $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:currentCsvPath)
        $sfd.FileName = [System.IO.Path]::GetFileName($script:currentCsvPath)
    }

    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        # Renumber No column (10, 20, 30...)
        $counter = 0
        $exportData = @()
        foreach ($row in $dt.Rows) {
            $counter += 10
            $exportData += [PSCustomObject]@{
                No          = $counter
                AppName     = $row["AppName"]
                Enabled     = $row["Enabled"]
                Description = $row["Description"]
            }
        }

        $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding Default
        $script:currentCsvPath = $sfd.FileName
        $lblCsvPath.Text = $sfd.FileName
        $script:isDirty = $false

        # Refresh No in DataTable to match saved values
        $counter = 0
        foreach ($row in $dt.Rows) {
            $counter += 10
            $row["No"] = $counter.ToString()
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Saved: $($sfd.FileName)",
            "Saved",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save CSV: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

# --- Mark dirty on CSV cell edit ---
$gridCsv.Add_CellValueChanged({
    $script:isDirty = $true
})

# --- Form Closing: Unsaved check ---
$form.Add_FormClosing({
    param($sender, $e)
    if ($script:isDirty) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "You have unsaved changes. Close without saving?",
            "Confirm Close",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $e.Cancel = $true
        }
    }
})

# ========================================
# Show Form
# ========================================
$form.ShowDialog() | Out-Null
$form.Dispose()
