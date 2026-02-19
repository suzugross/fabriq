# ========================================
# Bloatware Exporter GUI for Fabriq
# ========================================
# GUI tool for scanning installed legacy desktop
# applications from the registry and editing
# bloatware_list.csv for the Bloatware Remove module.
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
$fgWarning    = [System.Drawing.Color]::FromArgb(255, 160, 50)

# ========================================
# Fixed CSV Path (bloatware_remove module)
# ========================================
$script:csvPath = Join-Path $PSScriptRoot "..\..\modules\standard\bloatware_remove\bloatware_list.csv"
$script:csvPath = [System.IO.Path]::GetFullPath($script:csvPath)

# ========================================
# State
# ========================================
$script:isDirty      = $false
$script:allScanItems = @()

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

    $csvData = Import-Csv -Path $Path -Encoding UTF8
    $Table.Rows.Clear()

    $columns = @(
        "Enabled", "DisplayName", "Publisher", "DisplayVersion", "Architecture",
        "WindowsInstaller", "QuietUninstallString", "UninstallString",
        "NoRemove", "SystemComponent", "InstallDate", "RegistryKey"
    )

    foreach ($item in $csvData) {
        $props = $item.PSObject.Properties.Name
        $row = $Table.NewRow()
        foreach ($col in $columns) {
            if ($col -in $props) { $row[$col] = $item.$col } else { $row[$col] = "" }
        }
        if ([string]::IsNullOrWhiteSpace($row["Enabled"])) { $row["Enabled"] = "0" }
        $Table.Rows.Add($row)
    }
}

# ========================================
# Helper: Filter scan results grid
# ========================================
function Update-ScanFilter {
    param([string]$FilterText)

    $gridScan.Rows.Clear()
    $filter = $FilterText.Trim().ToLower()

    foreach ($item in $script:allScanItems) {
        if ($filter -and
            -not ($item.DisplayName.ToLower().Contains($filter)) -and
            -not ($item.Publisher.ToLower().Contains($filter))) {
            continue
        }
        $null = $gridScan.Rows.Add(
            $item.DisplayName,
            $item.Publisher,
            $item.DisplayVersion,
            $item.Architecture,
            $item.WindowsInstaller,
            $item.NoRemove,
            $item.SystemComponent
        )
    }

    $lblScanCount.Text = "Showing: $($gridScan.Rows.Count) / $($script:allScanItems.Count)"
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Bloatware Exporter - Fabriq"
$form.Size = New-Object System.Drawing.Size(1150, 800)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgDark
$form.ForeColor = $fgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ========================================
# SplitContainer (Top: Scanner, Bottom: CSV)
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
# TOP PANEL: Registry Scanner
# ==========================================
$panelTop = $splitContainer.Panel1
$panelTop.Padding = New-Object System.Windows.Forms.Padding(10)

# --- Header ---
$lblTopHeader = New-Object System.Windows.Forms.Label
$lblTopHeader.Text = "Installed Apps (Registry)"
$lblTopHeader.Location = New-Object System.Drawing.Point(10, 8)
$lblTopHeader.Size = New-Object System.Drawing.Size(220, 20)
$lblTopHeader.ForeColor = $fgHeader
$lblTopHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelTop.Controls.Add($lblTopHeader)

# --- Filter + Action Row ---
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Location = New-Object System.Drawing.Point(10, 36)
$lblFilter.Size = New-Object System.Drawing.Size(45, 20)
$lblFilter.ForeColor = $fgText
$panelTop.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(58, 34)
$txtFilter.Size = New-Object System.Drawing.Size(310, 25)
Set-TextBoxStyle $txtFilter
$panelTop.Controls.Add($txtFilter)

$btnScan = New-StyledButton -Text "Scan Registry" -X 378 -Y 32 -Width 120 -BgColor $bgAccent
$panelTop.Controls.Add($btnScan)

$btnAddToCsv = New-StyledButton -Text "Add to CSV >>" -X 508 -Y 32 -Width 130 -BgColor $bgAdd
$panelTop.Controls.Add($btnAddToCsv)

# --- Warning note ---
$lblWarningNote = New-Object System.Windows.Forms.Label
$lblWarningNote.Text = "* Orange = NoRemove=1 or SystemComponent=1 (handle with care)"
$lblWarningNote.Location = New-Object System.Drawing.Point(650, 36)
$lblWarningNote.Size = New-Object System.Drawing.Size(450, 18)
$lblWarningNote.ForeColor = $fgWarning
$lblWarningNote.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$panelTop.Controls.Add($lblWarningNote)

# --- Scan count ---
$lblScanCount = New-Object System.Windows.Forms.Label
$lblScanCount.Location = New-Object System.Drawing.Point(10, 60)
$lblScanCount.Size = New-Object System.Drawing.Size(500, 18)
$lblScanCount.Text = "(Not yet scanned — click 'Scan Registry')"
$lblScanCount.ForeColor = $fgDim
$panelTop.Controls.Add($lblScanCount)

# --- Scan Grid ---
$gridScan = New-Object System.Windows.Forms.DataGridView
$gridScan.Location = New-Object System.Drawing.Point(10, 80)
$gridScan.Size = New-Object System.Drawing.Size(1110, 270)
$gridScan.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$gridScan.SelectionMode = "FullRowSelect"
$gridScan.MultiSelect = $true
$gridScan.ReadOnly = $true
$gridScan.AutoSizeColumnsMode = "Fill"
Set-GridStyle $gridScan

$colSDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSDisplayName.Name = "DisplayName"
$colSDisplayName.HeaderText = "DisplayName"
$colSDisplayName.FillWeight = 30
$null = $gridScan.Columns.Add($colSDisplayName)

$colSPublisher = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSPublisher.Name = "Publisher"
$colSPublisher.HeaderText = "Publisher"
$colSPublisher.FillWeight = 20
$null = $gridScan.Columns.Add($colSPublisher)

$colSVersion = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSVersion.Name = "DisplayVersion"
$colSVersion.HeaderText = "Version"
$colSVersion.FillWeight = 10
$null = $gridScan.Columns.Add($colSVersion)

$colSArch = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSArch.Name = "Architecture"
$colSArch.HeaderText = "Arch"
$colSArch.FillWeight = 7
$null = $gridScan.Columns.Add($colSArch)

$colSWI = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSWI.Name = "WindowsInstaller"
$colSWI.HeaderText = "MSI"
$colSWI.FillWeight = 5
$null = $gridScan.Columns.Add($colSWI)

$colSNoRemove = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSNoRemove.Name = "NoRemove"
$colSNoRemove.HeaderText = "NoRem"
$colSNoRemove.FillWeight = 5
$null = $gridScan.Columns.Add($colSNoRemove)

$colSSysComp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSSysComp.Name = "SystemComponent"
$colSSysComp.HeaderText = "SysComp"
$colSSysComp.FillWeight = 5
$null = $gridScan.Columns.Add($colSSysComp)

$panelTop.Controls.Add($gridScan)

# ==========================================
# BOTTOM PANEL: CSV Editor
# ==========================================
$panelBottom = $splitContainer.Panel2
$panelBottom.Padding = New-Object System.Windows.Forms.Padding(10)

# --- CSV Toolbar ---
$lblCsvHeader = New-Object System.Windows.Forms.Label
$lblCsvHeader.Text = "bloatware_list.csv"
$lblCsvHeader.Location = New-Object System.Drawing.Point(10, 10)
$lblCsvHeader.Size = New-Object System.Drawing.Size(150, 20)
$lblCsvHeader.ForeColor = $fgHeader
$lblCsvHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelBottom.Controls.Add($lblCsvHeader)

$btnLoad      = New-StyledButton -Text "Load CSV"   -X 170 -Y 6 -Width 90 -BgColor $bgAccent
$btnSave      = New-StyledButton -Text "Save CSV"   -X 270 -Y 6 -Width 90 -BgColor $bgAccent
$btnDeleteRow = New-StyledButton -Text "Delete Row" -X 370 -Y 6 -Width 100 -BgColor $bgDelete

$panelBottom.Controls.Add($btnLoad)
$panelBottom.Controls.Add($btnSave)
$panelBottom.Controls.Add($btnDeleteRow)

# --- Path label ---
$lblCsvPath = New-Object System.Windows.Forms.Label
$lblCsvPath.Location = New-Object System.Drawing.Point(480, 10)
$lblCsvPath.Size = New-Object System.Drawing.Size(630, 18)
$lblCsvPath.Text = $script:csvPath
$lblCsvPath.ForeColor = $fgDim
$lblCsvPath.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$panelBottom.Controls.Add($lblCsvPath)

# --- CSV DataTable (12 columns) ---
$dt = New-Object System.Data.DataTable
foreach ($col in @(
    "Enabled", "DisplayName", "Publisher", "DisplayVersion", "Architecture",
    "WindowsInstaller", "QuietUninstallString", "UninstallString",
    "NoRemove", "SystemComponent", "InstallDate", "RegistryKey"
)) {
    [void]$dt.Columns.Add($col)
}

# --- CSV Grid ---
$gridCsv = New-Object System.Windows.Forms.DataGridView
$gridCsv.Location = New-Object System.Drawing.Point(10, 38)
$gridCsv.Size = New-Object System.Drawing.Size(1110, 280)
$gridCsv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$gridCsv.AutoSizeColumnsMode = "Fill"
$gridCsv.SelectionMode = "FullRowSelect"
$gridCsv.MultiSelect = $true
Set-GridStyle $gridCsv
$gridCsv.AllowUserToAddRows = $false
$gridCsv.AutoGenerateColumns = $false
$gridCsv.DataSource = $dt

# Enabled (ComboBox: 0/1)
$colEnabled = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colEnabled.HeaderText = "Enabled"
$colEnabled.DataPropertyName = "Enabled"
$colEnabled.Name = "Enabled"
$colEnabled.Items.AddRange("0", "1")
$colEnabled.FlatStyle = "Flat"
$colEnabled.FillWeight = 4
$colEnabled.MinimumWidth = 55
$null = $gridCsv.Columns.Add($colEnabled)

# DisplayName
$colDN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDN.HeaderText = "DisplayName"
$colDN.DataPropertyName = "DisplayName"
$colDN.Name = "DisplayName"
$colDN.FillWeight = 20
$null = $gridCsv.Columns.Add($colDN)

# Publisher
$colPub = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPub.HeaderText = "Publisher"
$colPub.DataPropertyName = "Publisher"
$colPub.Name = "Publisher"
$colPub.FillWeight = 12
$null = $gridCsv.Columns.Add($colPub)

# DisplayVersion
$colVer = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colVer.HeaderText = "Version"
$colVer.DataPropertyName = "DisplayVersion"
$colVer.Name = "DisplayVersion"
$colVer.FillWeight = 6
$null = $gridCsv.Columns.Add($colVer)

# Architecture
$colArch = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colArch.HeaderText = "Arch"
$colArch.DataPropertyName = "Architecture"
$colArch.Name = "Architecture"
$colArch.FillWeight = 5
$null = $gridCsv.Columns.Add($colArch)

# WindowsInstaller
$colWI = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colWI.HeaderText = "MSI"
$colWI.DataPropertyName = "WindowsInstaller"
$colWI.Name = "WindowsInstaller"
$colWI.FillWeight = 4
$null = $gridCsv.Columns.Add($colWI)

# QuietUninstallString
$colQUS = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colQUS.HeaderText = "QuietUninstall"
$colQUS.DataPropertyName = "QuietUninstallString"
$colQUS.Name = "QuietUninstallString"
$colQUS.FillWeight = 18
$null = $gridCsv.Columns.Add($colQUS)

# UninstallString
$colUS = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUS.HeaderText = "UninstallString"
$colUS.DataPropertyName = "UninstallString"
$colUS.Name = "UninstallString"
$colUS.FillWeight = 18
$null = $gridCsv.Columns.Add($colUS)

# NoRemove
$colNR = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNR.HeaderText = "NoRem"
$colNR.DataPropertyName = "NoRemove"
$colNR.Name = "NoRemove"
$colNR.FillWeight = 4
$null = $gridCsv.Columns.Add($colNR)

# SystemComponent
$colSC = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSC.HeaderText = "SysComp"
$colSC.DataPropertyName = "SystemComponent"
$colSC.Name = "SystemComponent"
$colSC.FillWeight = 5
$null = $gridCsv.Columns.Add($colSC)

# InstallDate
$colDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDate.HeaderText = "InstallDate"
$colDate.DataPropertyName = "InstallDate"
$colDate.Name = "InstallDate"
$colDate.FillWeight = 7
$null = $gridCsv.Columns.Add($colDate)

# RegistryKey
$colRK = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRK.HeaderText = "RegistryKey"
$colRK.DataPropertyName = "RegistryKey"
$colRK.Name = "RegistryKey"
$colRK.FillWeight = 18
$null = $gridCsv.Columns.Add($colRK)

$panelBottom.Controls.Add($gridCsv)

# ========================================
# Assemble Layout
# ========================================
$form.Controls.Add($splitContainer)

# ========================================
# Events
# ========================================

# --- Form Load: Auto-load default CSV ---
$form.Add_Load({
    if (Test-Path $script:csvPath) {
        try {
            Load-CsvToTable -Path $script:csvPath -Table $dt
            $script:isDirty = $false
        }
        catch {
            $lblCsvPath.Text = "(Failed to load: $($_.Exception.Message))"
        }
    }
    else {
        $lblCsvPath.Text = "$($script:csvPath)  [not found — will be created on Save]"
    }
})

# --- Scan Registry ---
$btnScan.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblScanCount.Text = "Scanning..."
    $gridScan.Rows.Clear()
    $script:allScanItems = @()

    $regPaths = @(
        [PSCustomObject]@{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*";             Arch = "64bit" }
        [PSCustomObject]@{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"; Arch = "32bit" }
    )

    $seenNames = @{}

    try {
        foreach ($rp in $regPaths) {
            $entries = @(Get-ItemProperty $rp.Path -ErrorAction SilentlyContinue)
            foreach ($entry in $entries) {
                if ([string]::IsNullOrWhiteSpace($entry.DisplayName)) { continue }
                if ($seenNames.ContainsKey($entry.DisplayName))       { continue }
                $seenNames[$entry.DisplayName] = $true

                $script:allScanItems += [PSCustomObject]@{
                    Enabled              = "0"
                    DisplayName          = $entry.DisplayName
                    Publisher            = if ($entry.Publisher)            { $entry.Publisher }            else { "" }
                    DisplayVersion       = if ($entry.DisplayVersion)       { $entry.DisplayVersion }       else { "" }
                    Architecture         = $rp.Arch
                    WindowsInstaller     = if ($entry.WindowsInstaller)     { "$($entry.WindowsInstaller)" } else { "0" }
                    QuietUninstallString = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { "" }
                    UninstallString      = if ($entry.UninstallString)      { $entry.UninstallString }      else { "" }
                    NoRemove             = if ($entry.NoRemove)             { "$($entry.NoRemove)" }        else { "0" }
                    SystemComponent      = if ($entry.SystemComponent)      { "$($entry.SystemComponent)" } else { "0" }
                    InstallDate          = if ($entry.InstallDate)          { $entry.InstallDate }          else { "" }
                    RegistryKey          = $entry.PSPath `
                                            -replace "Microsoft.PowerShell.Core\\Registry::HKEY_LOCAL_MACHINE", "HKLM:" `
                                            -replace "Microsoft.PowerShell.Core\\Registry::HKEY_CURRENT_USER",  "HKCU:"
                }
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Scan error: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }

    $script:allScanItems = @($script:allScanItems | Sort-Object Publisher, DisplayName)
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Update-ScanFilter -FilterText $txtFilter.Text
})

# --- Filter TextBox: real-time ---
$txtFilter.Add_TextChanged({
    Update-ScanFilter -FilterText $txtFilter.Text
})

# --- Scan grid: orange highlight for flagged rows ---
$gridScan.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $noRem = $gridScan.Rows[$e.RowIndex].Cells["NoRemove"].Value
    $sysCo = $gridScan.Rows[$e.RowIndex].Cells["SystemComponent"].Value
    if ($noRem -eq "1" -or $sysCo -eq "1") {
        $e.CellStyle.ForeColor = $fgWarning
    }
})

# --- Add to CSV ---
$btnAddToCsv.Add_Click({
    if ($gridScan.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select app(s) from the scan results first.",
            "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $addedCount   = 0
    $skippedCount = 0

    foreach ($selRow in $gridScan.SelectedRows) {
        $displayName = $selRow.Cells["DisplayName"].Value
        if ([string]::IsNullOrWhiteSpace($displayName)) { continue }

        # Duplicate check by DisplayName
        $existing = $dt.Select("DisplayName = '$($displayName -replace "'","''")'")
        if ($existing.Count -gt 0) {
            $skippedCount++
            continue
        }

        # Retrieve full data from cached scan list
        $src = $script:allScanItems | Where-Object { $_.DisplayName -eq $displayName } | Select-Object -First 1
        if ($null -eq $src) { continue }

        $row = $dt.NewRow()
        $row["Enabled"]              = "0"
        $row["DisplayName"]          = $src.DisplayName
        $row["Publisher"]            = $src.Publisher
        $row["DisplayVersion"]       = $src.DisplayVersion
        $row["Architecture"]         = $src.Architecture
        $row["WindowsInstaller"]     = $src.WindowsInstaller
        $row["QuietUninstallString"] = $src.QuietUninstallString
        $row["UninstallString"]      = $src.UninstallString
        $row["NoRemove"]             = $src.NoRemove
        $row["SystemComponent"]      = $src.SystemComponent
        $row["InstallDate"]          = $src.InstallDate
        $row["RegistryKey"]          = $src.RegistryKey
        $dt.Rows.Add($row)
        $addedCount++
    }

    if ($addedCount -gt 0) { $script:isDirty = $true }

    if ($skippedCount -gt 0) {
        $lblScanCount.Text = "Added: $addedCount, Skipped (duplicate): $skippedCount"
    }
    elseif ($addedCount -gt 0) {
        $lblScanCount.Text = "Added: $addedCount row(s) to CSV editor"
    }
})

# --- Delete Row ---
$btnDeleteRow.Add_Click({
    if ($gridCsv.SelectedRows.Count -eq 0) { return }

    $indices = @()
    foreach ($row in $gridCsv.SelectedRows) {
        if ($row.Index -lt $dt.Rows.Count) { $indices += $row.Index }
    }
    $indices = $indices | Sort-Object -Descending

    foreach ($idx in $indices) {
        $dt.Rows.RemoveAt($idx)
    }
    $script:isDirty = $true
})

# --- Load CSV ---
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
    $ofd.Title = "Load bloatware_list.csv"
    $ofd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:csvPath)

    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        Load-CsvToTable -Path $ofd.FileName -Table $dt
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

# --- Save CSV (always writes to fixed path) ---
$btnSave.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Save $($dt.Rows.Count) row(s) to:`n$($script:csvPath)?",
        "Confirm Save",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        $saveDir = [System.IO.Path]::GetDirectoryName($script:csvPath)
        if (-not (Test-Path $saveDir)) {
            $null = New-Item -ItemType Directory -Path $saveDir -Force
        }

        $exportData = @()
        foreach ($row in $dt.Rows) {
            $exportData += [PSCustomObject]@{
                Enabled              = $row["Enabled"]
                DisplayName          = $row["DisplayName"]
                Publisher            = $row["Publisher"]
                DisplayVersion       = $row["DisplayVersion"]
                Architecture         = $row["Architecture"]
                WindowsInstaller     = $row["WindowsInstaller"]
                QuietUninstallString = $row["QuietUninstallString"]
                UninstallString      = $row["UninstallString"]
                NoRemove             = $row["NoRemove"]
                SystemComponent      = $row["SystemComponent"]
                InstallDate          = $row["InstallDate"]
                RegistryKey          = $row["RegistryKey"]
            }
        }

        $exportData | Export-Csv -Path $script:csvPath -NoTypeInformation -Encoding UTF8 -Force
        $lblCsvPath.Text = $script:csvPath
        $script:isDirty = $false

        [System.Windows.Forms.MessageBox]::Show(
            "Saved $($dt.Rows.Count) row(s) to:`n$($script:csvPath)",
            "Saved",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save: $($_.Exception.Message)",
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

# --- Form Closing: unsaved check ---
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
