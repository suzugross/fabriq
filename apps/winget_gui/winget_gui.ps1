# ========================================
# Winget GUI Editor for Fabriq
# ========================================
# GUI tool for searching winget packages,
# viewing details, and editing app_list.csv
# for the Winget App Install module.
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
$script:defaultCsvPath = Join-Path $PSScriptRoot "..\..\modules\extended\winget_install\app_list.csv"
$script:defaultCsvPath = [System.IO.Path]::GetFullPath($script:defaultCsvPath)

# ========================================
# State Tracking
# ========================================
$script:isDirty = $false
$script:currentCsvPath = $null
$script:runspace = $null
$script:asyncHandle = $null

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
# Helper: Parse winget output (fixed-width)
# ========================================
function Parse-WingetOutput {
    param([string]$Output)

    $results = @()
    $lines = $Output -split "`r?`n"

    # Find header line
    $headerIdx = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^Name\s+Id\s+Version") {
            $headerIdx = $i
            break
        }
    }

    if ($headerIdx -lt 0) { return $results }

    $headerLine = $lines[$headerIdx]

    # Detect column start positions from header
    $idStart      = $headerLine.IndexOf("Id")
    $versionStart = $headerLine.IndexOf("Version")
    $sourceStart  = $headerLine.IndexOf("Source")

    if ($idStart -lt 0 -or $versionStart -lt 0) { return $results }

    # Skip header + separator line
    $dataStart = $headerIdx + 1
    if ($dataStart -lt $lines.Count -and $lines[$dataStart] -match "^-+") {
        $dataStart++
    }

    for ($i = $dataStart; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.Length -lt $versionStart) { continue }

        $name = $line.Substring(0, [Math]::Min($idStart, $line.Length)).TrimEnd()
        $idEnd = [Math]::Min($versionStart, $line.Length)
        $id = $line.Substring($idStart, $idEnd - $idStart).TrimEnd()

        $version = ""
        if ($line.Length -gt $versionStart) {
            if ($sourceStart -gt 0 -and $line.Length -gt $sourceStart) {
                $version = $line.Substring($versionStart, $sourceStart - $versionStart).TrimEnd()
            }
            else {
                $version = $line.Substring($versionStart).TrimEnd()
            }
        }

        # Accept any ID with length > 2 (includes Store IDs without dots)
        if ($id.Length -gt 2) {
            $results += [PSCustomObject]@{
                Name    = $name
                ID      = $id
                Version = $version
            }
        }
    }

    return $results
}

# ========================================
# Helper: Run winget command synchronously
# ========================================
function Invoke-Winget {
    param([string]$Arguments)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "winget.exe"
    $psi.Arguments = $Arguments
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.CreateNoWindow = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $output = $proc.StandardOutput.ReadToEnd()
    $proc.WaitForExit()
    return $output
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

        # Support both Fabriq format (AppID) and legacy format (Id)
        $enabled = ""
        $appId   = ""
        $options = ""
        $desc    = ""

        if ("Enabled" -in $props)     { $enabled = $item.Enabled }
        if ("AppID" -in $props)       { $appId = $item.AppID }
        elseif ("Id" -in $props)      { $appId = $item.Id }
        if ("Options" -in $props)     { $options = $item.Options }
        if ("Description" -in $props) { $desc = $item.Description }
        elseif ("Name" -in $props)    { $desc = $item.Name }

        # Default Enabled to 1 if missing
        if ([string]::IsNullOrWhiteSpace($enabled)) { $enabled = "1" }

        $Table.Rows.Add($enabled, $appId, $options, $desc)
    }
}

# ========================================
# Main Form
# ========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Winget GUI Editor - Fabriq"
$form.Size = New-Object System.Drawing.Size(1050, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor = $bgDark
$form.ForeColor = $fgText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ========================================
# SplitContainer (Top: Search, Bottom: CSV)
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
# TOP PANEL: Winget Search
# ==========================================
$panelTop = $splitContainer.Panel1
$panelTop.Padding = New-Object System.Windows.Forms.Padding(10)

# --- Network Status ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10, 8)
$lblStatus.Size = New-Object System.Drawing.Size(400, 18)
$lblStatus.Text = "Checking network..."
$lblStatus.ForeColor = $fgDim
$panelTop.Controls.Add($lblStatus)

# --- Search Row ---
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.Location = New-Object System.Drawing.Point(10, 36)
$lblSearch.Size = New-Object System.Drawing.Size(55, 20)
$lblSearch.ForeColor = $fgText
$panelTop.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(68, 34)
$txtSearch.Size = New-Object System.Drawing.Size(350, 25)
Set-TextBoxStyle $txtSearch
$panelTop.Controls.Add($txtSearch)

$btnSearch = New-StyledButton -Text "Search" -X 425 -Y 32 -Width 90 -BgColor $bgAccent
$panelTop.Controls.Add($btnSearch)

$btnShowDetails = New-StyledButton -Text "Show Details" -X 525 -Y 32 -Width 110
$panelTop.Controls.Add($btnShowDetails)

# --- Search Results Grid ---
$gridSearch = New-Object System.Windows.Forms.DataGridView
$gridSearch.Location = New-Object System.Drawing.Point(10, 65)
$gridSearch.Size = New-Object System.Drawing.Size(1010, 230)
$gridSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$gridSearch.SelectionMode = "FullRowSelect"
$gridSearch.MultiSelect = $true
$gridSearch.ReadOnly = $true
$gridSearch.AutoSizeColumnsMode = "Fill"
Set-GridStyle $gridSearch

$colSName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSName.Name = "Name"
$colSName.HeaderText = "Name"
$colSName.FillWeight = 35
$null = $gridSearch.Columns.Add($colSName)

$colSID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSID.Name = "ID"
$colSID.HeaderText = "ID"
$colSID.FillWeight = 35
$null = $gridSearch.Columns.Add($colSID)

$colSVer = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSVer.Name = "Version"
$colSVer.HeaderText = "Version"
$colSVer.FillWeight = 30
$null = $gridSearch.Columns.Add($colSVer)

$panelTop.Controls.Add($gridSearch)

# --- Options + Add to CSV Row ---
$lblOptions = New-Object System.Windows.Forms.Label
$lblOptions.Text = "Options:"
$lblOptions.Location = New-Object System.Drawing.Point(10, 305)
$lblOptions.Size = New-Object System.Drawing.Size(60, 20)
$lblOptions.ForeColor = $fgText
$lblOptions.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$panelTop.Controls.Add($lblOptions)

$txtOptions = New-Object System.Windows.Forms.TextBox
$txtOptions.Location = New-Object System.Drawing.Point(68, 303)
$txtOptions.Size = New-Object System.Drawing.Size(350, 25)
$txtOptions.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
Set-TextBoxStyle $txtOptions
$panelTop.Controls.Add($txtOptions)

$btnAdd = New-StyledButton -Text "Add to CSV" -X 425 -Y 301 -Width 110 -BgColor $bgAdd
$btnAdd.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$panelTop.Controls.Add($btnAdd)

# --- Search result count label ---
$lblSearchCount = New-Object System.Windows.Forms.Label
$lblSearchCount.Location = New-Object System.Drawing.Point(650, 36)
$lblSearchCount.Size = New-Object System.Drawing.Size(350, 18)
$lblSearchCount.Text = ""
$lblSearchCount.ForeColor = $fgDim
$lblSearchCount.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$panelTop.Controls.Add($lblSearchCount)

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
[void]$dt.Columns.Add("Enabled")
[void]$dt.Columns.Add("AppID")
[void]$dt.Columns.Add("Options")
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
$gridCsv.DataSource = $dt

$panelBottom.Controls.Add($gridCsv)

# ========================================
# Assemble Layout
# ========================================
$form.Controls.Add($splitContainer)

# ========================================
# Async Search Timer
# ========================================
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200

$timer.Add_Tick({
    if ($null -ne $script:asyncHandle -and $script:asyncHandle.IsCompleted) {
        $timer.Stop()

        try {
            $output = $script:runspace.EndInvoke($script:asyncHandle)
            $rawText = $output -join "`r`n"

            $parsed = Parse-WingetOutput -Output $rawText
            $gridSearch.Rows.Clear()

            foreach ($item in $parsed) {
                $null = $gridSearch.Rows.Add($item.Name, $item.ID, $item.Version)
            }

            $lblSearchCount.Text = "Results: $($parsed.Count)"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Search failed: $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            $lblSearchCount.Text = "Search failed"
        }
        finally {
            if ($null -ne $script:runspace) {
                $script:runspace.Dispose()
                $script:runspace = $null
            }
            $script:asyncHandle = $null
            $btnSearch.Enabled = $true
            $btnSearch.Text = "Search"
        }
    }
})

# ========================================
# Events
# ========================================

# --- Form Load: Network Check + Default CSV ---
$form.Add_Load({
    # Network check
    try {
        $ping = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
        if ($ping) {
            $lblStatus.Text = "Network: OK (8.8.8.8)"
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(80, 200, 80)
        }
        else {
            $lblStatus.Text = "Network: FAILED (8.8.8.8 unreachable)"
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(255, 80, 80)
        }
    }
    catch {
        $lblStatus.Text = "Network: Error"
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(255, 80, 80)
    }

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

# --- Search Button: Async winget search ---
$btnSearch.Add_Click({
    $query = $txtSearch.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($query)) { return }

    $gridSearch.Rows.Clear()
    $lblSearchCount.Text = "Searching..."
    $btnSearch.Enabled = $false
    $btnSearch.Text = "Searching..."

    # Create Runspace for async execution
    $script:runspace = [PowerShell]::Create()
    $null = $script:runspace.AddScript({
        param($Query)
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "winget.exe"
        $psi.Arguments = "search `"$Query`" --accept-source-agreements"
        $psi.RedirectStandardOutput = $true
        $psi.UseShellExecute = $false
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.CreateNoWindow = $true

        $proc = [System.Diagnostics.Process]::Start($psi)
        $output = $proc.StandardOutput.ReadToEnd()
        $proc.WaitForExit()
        return $output
    })
    $null = $script:runspace.AddArgument($query)

    $script:asyncHandle = $script:runspace.BeginInvoke()
    $timer.Start()
})

# --- Enter key triggers search ---
$txtSearch.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $btnSearch.PerformClick()
    }
})

# --- Show Details Button ---
$btnShowDetails.Add_Click({
    if ($gridSearch.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select an app from the search results first.",
            "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $selectedId = $gridSearch.SelectedRows[0].Cells["ID"].Value
    if ([string]::IsNullOrWhiteSpace($selectedId)) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        $details = Invoke-Winget -Arguments "show `"$selectedId`" --exact --accept-source-agreements"

        # Show in a details dialog
        $detailForm = New-Object System.Windows.Forms.Form
        $detailForm.Text = "Package Details - $selectedId"
        $detailForm.Size = New-Object System.Drawing.Size(650, 500)
        $detailForm.StartPosition = "CenterParent"
        $detailForm.BackColor = $bgDark
        $detailForm.ForeColor = $fgText
        $detailForm.Font = New-Object System.Drawing.Font("Consolas", 9)

        $txtDetail = New-Object System.Windows.Forms.TextBox
        $txtDetail.Dock = "Fill"
        $txtDetail.Multiline = $true
        $txtDetail.ReadOnly = $true
        $txtDetail.ScrollBars = "Both"
        $txtDetail.WordWrap = $false
        $txtDetail.BackColor = $bgCell
        $txtDetail.ForeColor = $fgText
        $txtDetail.Text = $details

        $detailForm.Controls.Add($txtDetail)
        $detailForm.ShowDialog() | Out-Null
        $detailForm.Dispose()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to get details: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# --- Add to CSV Button ---
$btnAdd.Add_Click({
    if ($gridSearch.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select app(s) from the search results first.",
            "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $opts = $txtOptions.Text
    $addedCount = 0
    $skippedCount = 0

    foreach ($selRow in $gridSearch.SelectedRows) {
        $id   = $selRow.Cells["ID"].Value
        $name = $selRow.Cells["Name"].Value

        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        # Duplicate check
        $existing = $dt.Select("AppID = '$($id -replace "'","''")'")
        if ($existing.Count -gt 0) {
            $skippedCount++
            continue
        }

        $dt.Rows.Add("1", $id, $opts, $name)
        $addedCount++
    }

    if ($addedCount -gt 0) {
        $script:isDirty = $true
    }

    if ($skippedCount -gt 0) {
        $lblSearchCount.Text = "Added: $addedCount, Skipped (duplicate): $skippedCount"
    }
    else {
        $lblSearchCount.Text = "Added: $addedCount"
    }
})

# --- Delete Row Button ---
$btnDeleteRow.Add_Click({
    if ($gridCsv.SelectedRows.Count -eq 0) { return }

    # Collect indices in descending order to avoid index shift
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

    # Default to winget_install module directory
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
    $sfd.FileName = "app_list.csv"

    # Default to last loaded path
    if ($script:currentCsvPath) {
        $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:currentCsvPath)
        $sfd.FileName = [System.IO.Path]::GetFileName($script:currentCsvPath)
    }

    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $exportData = @()
        foreach ($row in $dt.Rows) {
            $exportData += [PSCustomObject]@{
                Enabled     = $row["Enabled"]
                AppID       = $row["AppID"]
                Options     = $row["Options"]
                Description = $row["Description"]
            }
        }

        $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding Default
        $script:currentCsvPath = $sfd.FileName
        $lblCsvPath.Text = $sfd.FileName
        $script:isDirty = $false

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

    # Cleanup async resources
    if ($null -ne $script:runspace) {
        $script:runspace.Dispose()
        $script:runspace = $null
    }
})

# ========================================
# Show Form
# ========================================
$form.ShowDialog() | Out-Null
$form.Dispose()
