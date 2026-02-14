# ========================================
# Fabriq CSV Editor
# ========================================
# GUI-based CSV editor accessible from Command Mode.
# Provides a ListView for CSV file selection and
# a DataGridView editor for editing CSV contents.
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========================================
# Dark Theme Colors (matching status_monitor / gyotaku)
# ========================================
$darkBg       = [System.Drawing.Color]::FromArgb(30, 30, 30)
$panelBg      = [System.Drawing.Color]::FromArgb(45, 45, 45)
$accentCyan   = [System.Drawing.Color]::FromArgb(0, 200, 200)
$textWhite    = [System.Drawing.Color]::White
$textGray     = [System.Drawing.Color]::FromArgb(160, 160, 160)
$successGreen = [System.Drawing.Color]::FromArgb(80, 220, 80)
$warnYellow   = [System.Drawing.Color]::FromArgb(255, 200, 0)
$errorRed     = [System.Drawing.Color]::FromArgb(255, 80, 80)
$btnGreenBg   = [System.Drawing.Color]::FromArgb(30, 100, 30)
$btnYellowBg  = [System.Drawing.Color]::FromArgb(100, 85, 20)
$btnRedBg     = [System.Drawing.Color]::FromArgb(100, 30, 30)
$btnCyanBg    = [System.Drawing.Color]::FromArgb(20, 60, 60)

$fontNormal = New-Object System.Drawing.Font("Consolas", 9)
$fontBold   = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$fontSmall  = New-Object System.Drawing.Font("Consolas", 9)

# ========================================
# CSV Registry: Editable CSV files
# ========================================
$script:CsvRegistry = @(
    @{ Name = "Host List";           Group = "Kernel";   Path = ".\kernel\hostlist.csv" }
    @{ Name = "Categories";          Group = "Kernel";   Path = ".\kernel\categories.csv" }
    @{ Name = "App Install List";    Group = "Apps";     Path = ".\modules\standard\app_config\app_list.csv" }
    @{ Name = "Registry HKLM";       Group = "Registry"; Path = ".\modules\standard\reg_config\reg_hklm_list.csv" }
    @{ Name = "Registry HKCU";       Group = "Registry"; Path = ".\modules\standard\reg_config\reg_hkcu_list.csv" }
    @{ Name = "Registry Firewall";   Group = "Registry"; Path = ".\modules\standard\reg_config\reg_hklm_list_firewall.csv" }
    @{ Name = "Registry Power";      Group = "Registry"; Path = ".\modules\standard\reg_config\reg_hklm_list_power.csv" }
    @{ Name = "Local Users";         Group = "Users";    Path = ".\modules\standard\local_user_config\local_user_list.csv" }
    @{ Name = "Power Profiles";      Group = "Power";    Path = ".\modules\standard\power_config\power_list.csv" }
    @{ Name = "Store Apps Remove";   Group = "Apps";     Path = ".\modules\standard\storeapp_config\storeapp_list.csv" }
    @{ Name = "Domain Settings";     Group = "Network";  Path = ".\modules\standard\domain_join\domain.csv" }
    @{ Name = "Gyotaku Tasks";       Group = "Evidence"; Path = ".\modules\standard\gyotaku_template\task_list.csv" }
    @{ Name = "IPv6 Config";         Group = "Network";  Path = ".\modules\extended\ipv6_config\ipv6_list.csv" }
    @{ Name = "Group Members";      Group = "Users";    Path = ".\modules\extended\group_config\group_list.csv" }
    @{ Name = "Display Resolution"; Group = "Registry"; Path = ".\modules\extended\display_config\display_list.csv" }
    @{ Name = "DPI Scaling";        Group = "Registry"; Path = ".\modules\extended\dpi_config\dpi_list.csv" }
    @{ Name = "License Keys";      Group = "System";   Path = ".\modules\standard\windows_license_config\license_key.csv" }
    @{ Name = "AutoKey Recipe";    Group = "Automation"; Path = ".\modules\standard\autokey_template\recipe.csv" }
    @{ Name = "File Copy List";   Group = "System";     Path = ".\modules\standard\copyfile_config\copy_list.csv" }
    @{ Name = "Reg Backup List";  Group = "Registry";   Path = ".\modules\extended\reg_template\reg_list.csv" }
)

# Dynamic Profile CSV discovery
$profileDir = ".\profiles"
if (Test-Path $profileDir) {
    $profileFiles = Get-ChildItem $profileDir -Filter "*.csv" -File -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($pf in $profileFiles) {
        $profileName = [System.IO.Path]::GetFileNameWithoutExtension($pf.Name)
        $script:CsvRegistry += @{ Name = "Profile: $profileName"; Group = "Profiles"; Path = $pf.FullName }
    }
}

# ========================================
# Helper Functions
# ========================================

function Detect-CsvEncoding {
    param([string]$FilePath)
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            return "UTF8"
        }
    } catch { }
    return "Default"
}

function Get-CsvInfo {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        return @{ Exists = $false; Rows = 0; Cols = 0; Encoding = "N/A"; Headers = @() }
    }
    try {
        $data = @(Import-Csv -Path $Path -Encoding Default)
        $cols = @()
        if ($data.Count -gt 0) {
            $cols = @($data[0].PSObject.Properties.Name)
        } else {
            # Header-only file: read first line manually
            $firstLine = [System.IO.File]::ReadLines($Path) | Select-Object -First 1
            if ($firstLine) {
                # Strip BOM if present
                $firstLine = $firstLine -replace '^\xEF\xBB\xBF', '' -replace '^\uFEFF', ''
                $cols = @($firstLine -split ',')
            }
        }
        return @{
            Exists   = $true
            Rows     = $data.Count
            Cols     = $cols.Count
            Encoding = (Detect-CsvEncoding -FilePath $Path)
            Headers  = $cols
        }
    }
    catch {
        return @{ Exists = $false; Rows = 0; Cols = 0; Encoding = "Error"; Headers = @() }
    }
}

# ========================================
# CSV Selection Form
# ========================================

function Show-CsvSelector {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Fabriq CSV Editor"
    $form.Size = New-Object System.Drawing.Size(750, 520)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.BackColor = $darkBg
    $form.ForeColor = $textWhite
    $form.Font = $fontNormal

    # --- Title Label ---
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Select a CSV file to edit"
    $lblTitle.Location = New-Object System.Drawing.Point(15, 12)
    $lblTitle.Size = New-Object System.Drawing.Size(700, 22)
    $lblTitle.Font = $fontBold
    $lblTitle.ForeColor = $accentCyan
    $null = $form.Controls.Add($lblTitle)

    # --- ListView ---
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(15, 40)
    $listView.Size = New-Object System.Drawing.Size(705, 340)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.HideSelection = $false
    $listView.MultiSelect = $false
    $listView.GridLines = $true
    $listView.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $listView.ForeColor = $textWhite
    $listView.Font = $fontNormal
    $listView.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable

    # OwnerDraw for custom header colors
    $listView.OwnerDraw = $true

    $listView.Add_DrawColumnHeader({
        param($sender, $e)
        $e.Graphics.FillRectangle(
            (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(50, 50, 50))),
            $e.Bounds
        )
        [System.Windows.Forms.TextRenderer]::DrawText(
            $e.Graphics,
            $e.Header.Text,
            $fontSmall,
            $e.Bounds,
            $accentCyan,
            [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::Left
        )
    })

    $listView.Add_DrawItem({
        param($sender, $e)
        $e.DrawDefault = $true
    })

    $listView.Add_DrawSubItem({
        param($sender, $e)
        $e.DrawDefault = $true
    })

    # Columns
    $null = $listView.Columns.Add("Group", 80)
    $null = $listView.Columns.Add("Name", 180)
    $null = $listView.Columns.Add("Rows", 55)
    $null = $listView.Columns.Add("Cols", 55)
    $null = $listView.Columns.Add("Encoding", 75)
    $null = $listView.Columns.Add("Path", 240)

    $null = $form.Controls.Add($listView)

    # --- Info Label ---
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(15, 388)
    $lblInfo.Size = New-Object System.Drawing.Size(705, 20)
    $lblInfo.Font = $fontSmall
    $lblInfo.ForeColor = $textGray
    $lblInfo.Text = "Select a CSV file from the list above"
    $null = $form.Controls.Add($lblInfo)

    # --- Buttons ---
    # Edit button
    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Location = New-Object System.Drawing.Point(15, 418)
    $btnEdit.Size = New-Object System.Drawing.Size(150, 38)
    $btnEdit.Text = "Edit"
    $btnEdit.Font = $fontBold
    $btnEdit.ForeColor = $successGreen
    $btnEdit.BackColor = $btnGreenBg
    $btnEdit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnEdit.FlatAppearance.BorderColor = $successGreen
    $btnEdit.FlatAppearance.BorderSize = 1
    $btnEdit.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnEdit.Enabled = $false
    $null = $form.Controls.Add($btnEdit)

    # Open File button
    $btnOpenFile = New-Object System.Windows.Forms.Button
    $btnOpenFile.Location = New-Object System.Drawing.Point(180, 418)
    $btnOpenFile.Size = New-Object System.Drawing.Size(150, 38)
    $btnOpenFile.Text = "Open File..."
    $btnOpenFile.Font = $fontSmall
    $btnOpenFile.ForeColor = $accentCyan
    $btnOpenFile.BackColor = $btnCyanBg
    $btnOpenFile.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpenFile.FlatAppearance.BorderColor = $accentCyan
    $btnOpenFile.FlatAppearance.BorderSize = 1
    $btnOpenFile.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $form.Controls.Add($btnOpenFile)

    # Refresh button
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Location = New-Object System.Drawing.Point(450, 418)
    $btnRefresh.Size = New-Object System.Drawing.Size(120, 38)
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Font = $fontSmall
    $btnRefresh.ForeColor = $warnYellow
    $btnRefresh.BackColor = $btnYellowBg
    $btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRefresh.FlatAppearance.BorderColor = $warnYellow
    $btnRefresh.FlatAppearance.BorderSize = 1
    $btnRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $form.Controls.Add($btnRefresh)

    # Close button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(585, 418)
    $btnClose.Size = New-Object System.Drawing.Size(135, 38)
    $btnClose.Text = "Close"
    $btnClose.Font = $fontSmall
    $btnClose.ForeColor = $errorRed
    $btnClose.BackColor = $btnRedBg
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.FlatAppearance.BorderColor = $errorRed
    $btnClose.FlatAppearance.BorderSize = 1
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $form.Controls.Add($btnClose)

    # --- Populate ListView ---
    function Populate-CsvList {
        $null = $listView.Items.Clear()
        foreach ($entry in $script:CsvRegistry) {
            $info = Get-CsvInfo -Path $entry.Path
            $item = New-Object System.Windows.Forms.ListViewItem($entry.Group)
            $null = $item.SubItems.Add($entry.Name)
            $null = $item.SubItems.Add($(if ($info.Exists) { $info.Rows.ToString() } else { "N/A" }))
            $null = $item.SubItems.Add($(if ($info.Exists) { $info.Cols.ToString() } else { "N/A" }))
            $null = $item.SubItems.Add($(if ($info.Exists) { $info.Encoding } else { "N/A" }))
            $null = $item.SubItems.Add($entry.Path)
            $item.Tag = $entry
            if (-not $info.Exists) {
                $item.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
            }
            $null = $listView.Items.Add($item)
        }
    }

    Populate-CsvList

    # --- Helper: Get selected entry ---
    function Get-SelectedEntry {
        if ($listView.SelectedItems.Count -eq 0) { return $null }
        return $listView.SelectedItems[0].Tag
    }

    # --- Helper: Open editor for a path ---
    function Open-EditorForPath {
        param([string]$CsvPath, [string]$CsvName)
        if (-not (Test-Path $CsvPath)) {
            $null = [System.Windows.Forms.MessageBox]::Show(
                "File not found: $CsvPath",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        $null = $form.Hide()
        Open-CsvEditor -CsvPath $CsvPath -CsvName $CsvName
        Populate-CsvList
        $null = $form.Show()
    }

    # --- Events ---
    $listView.Add_SelectedIndexChanged({
        $entry = Get-SelectedEntry
        if ($null -ne $entry) {
            $info = Get-CsvInfo -Path $entry.Path
            if ($info.Exists) {
                $lblInfo.Text = "$($entry.Name): $($info.Rows) rows x $($info.Cols) columns | $($info.Encoding) | $($entry.Path)"
                $lblInfo.ForeColor = $textWhite
                $btnEdit.Enabled = $true
            } else {
                $lblInfo.Text = "$($entry.Name): FILE NOT FOUND ($($entry.Path))"
                $lblInfo.ForeColor = $errorRed
                $btnEdit.Enabled = $false
            }
        } else {
            $lblInfo.Text = "Select a CSV file from the list above"
            $lblInfo.ForeColor = $textGray
            $btnEdit.Enabled = $false
        }
    })

    $listView.Add_DoubleClick({
        $entry = Get-SelectedEntry
        if ($null -ne $entry -and (Test-Path $entry.Path)) {
            Open-EditorForPath -CsvPath $entry.Path -CsvName $entry.Name
        }
    })

    $btnEdit.Add_Click({
        $entry = Get-SelectedEntry
        if ($null -ne $entry) {
            Open-EditorForPath -CsvPath $entry.Path -CsvName $entry.Name
        }
    })

    $btnOpenFile.Add_Click({
        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $openDialog.Title = "Open CSV File"
        if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($openDialog.FileName)
            Open-EditorForPath -CsvPath $openDialog.FileName -CsvName $fileName
        }
    })

    $btnRefresh.Add_Click({
        Populate-CsvList
        $lblInfo.Text = "List refreshed"
        $lblInfo.ForeColor = $successGreen
        $btnEdit.Enabled = $false
    })

    $btnClose.Add_Click({
        $form.Close()
    })

    # Show the selector form (modal)
    $null = $form.ShowDialog()
    $form.Dispose()
}

# ========================================
# CSV Editor Form (DataGridView)
# ========================================

function Open-CsvEditor {
    param(
        [string]$CsvPath,
        [string]$CsvName
    )

    # --- State ---
    $script:IsModified = $false
    $script:CsvEncoding = Detect-CsvEncoding -FilePath $CsvPath

    # --- Load CSV Data ---
    $csvData = $null
    try {
        $csvData = @(Import-Csv -Path $CsvPath -Encoding Default)
    }
    catch {
        $null = [System.Windows.Forms.MessageBox]::Show(
            "Failed to load CSV: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # Get column headers
    $headers = @()
    if ($csvData.Count -gt 0) {
        $headers = @($csvData[0].PSObject.Properties.Name)
    } else {
        $firstLine = [System.IO.File]::ReadLines($CsvPath) | Select-Object -First 1
        if ($firstLine) {
            $firstLine = $firstLine -replace '^\uFEFF', ''
            $headers = @($firstLine -split ',')
        }
    }

    if ($headers.Count -eq 0) {
        $null = [System.Windows.Forms.MessageBox]::Show(
            "CSV file has no columns",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # --- Build DataTable ---
    $dataTable = New-Object System.Data.DataTable
    foreach ($header in $headers) {
        $null = $dataTable.Columns.Add($header, [string])
    }
    foreach ($row in $csvData) {
        $newRow = $dataTable.NewRow()
        foreach ($header in $headers) {
            $newRow[$header] = if ($null -ne $row.$header) { $row.$header } else { "" }
        }
        $null = $dataTable.Rows.Add($newRow)
    }
    $dataTable.AcceptChanges()

    # --- Editor Form ---
    $formEditor = New-Object System.Windows.Forms.Form
    $formEditor.Text = "CSV Editor: $CsvName ($CsvPath)"
    $formEditor.Size = New-Object System.Drawing.Size(1200, 700)
    $formEditor.MinimumSize = New-Object System.Drawing.Size(800, 500)
    $formEditor.StartPosition = "CenterScreen"
    $formEditor.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $formEditor.BackColor = $darkBg
    $formEditor.ForeColor = $textWhite
    $formEditor.Font = $fontNormal
    $formEditor.KeyPreview = $true

    # --- Toolbar ---
    $toolPanel = New-Object System.Windows.Forms.Panel
    $toolPanel.Location = New-Object System.Drawing.Point(0, 0)
    $toolPanel.Size = New-Object System.Drawing.Size($formEditor.ClientSize.Width, 45)
    $toolPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $toolPanel.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $null = $formEditor.Controls.Add($toolPanel)

    # Helper: get current row index (-1 if none)
    $script:GetCurrentRowIndex = {
        if ($null -ne $dgv.CurrentRow) { return $dgv.CurrentRow.Index }
        return -1
    }

    # --- Row operation: Add ---
    $script:DoAddRow = {
        $insertIndex = & $script:GetCurrentRowIndex
        $newRow = $dataTable.NewRow()
        foreach ($h in $headers) { $newRow[$h] = "" }
        if ($insertIndex -ge 0 -and $insertIndex -lt $dataTable.Rows.Count) {
            $dataTable.Rows.InsertAt($newRow, $insertIndex + 1)
            $dgv.CurrentCell = $dgv.Rows[($insertIndex + 1)].Cells[0]
        } else {
            $null = $dataTable.Rows.Add($newRow)
            $dgv.CurrentCell = $dgv.Rows[($dataTable.Rows.Count - 1)].Cells[0]
        }
    }

    # --- Row operation: Copy ---
    $script:DoCopyRow = {
        $srcIndex = & $script:GetCurrentRowIndex
        if ($srcIndex -lt 0 -or $srcIndex -ge $dataTable.Rows.Count) { return }
        $srcRow = $dataTable.Rows[$srcIndex]
        $newRow = $dataTable.NewRow()
        foreach ($h in $headers) { $newRow[$h] = $srcRow[$h] }
        $dataTable.Rows.InsertAt($newRow, $srcIndex + 1)
        $dgv.CurrentCell = $dgv.Rows[($srcIndex + 1)].Cells[0]
    }

    # --- Row operation: Delete ---
    $script:DoDeleteRow = {
        # Collect unique row indices from selected cells
        $selectedIndices = @()
        foreach ($cell in $dgv.SelectedCells) {
            if ($selectedIndices -notcontains $cell.RowIndex) {
                $selectedIndices += $cell.RowIndex
            }
        }
        if ($selectedIndices.Count -eq 0) { return }

        $msg = if ($selectedIndices.Count -eq 1) {
            "Delete row $(($selectedIndices[0] + 1))?"
        } else {
            "Delete $($selectedIndices.Count) selected rows?"
        }
        $result = [System.Windows.Forms.MessageBox]::Show(
            $msg, "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        # Delete in reverse order to preserve indices
        $sorted = $selectedIndices | Sort-Object -Descending
        foreach ($idx in $sorted) {
            if ($idx -ge 0 -and $idx -lt $dataTable.Rows.Count) {
                $dataTable.Rows.RemoveAt($idx)
            }
        }
    }

    # --- Row operation: Move Up ---
    $script:DoMoveUp = {
        $null = $dgv.EndEdit()
        $idx = & $script:GetCurrentRowIndex
        if ($idx -le 0 -or $idx -ge $dataTable.Rows.Count) { return }
        foreach ($h in $headers) {
            $temp = $dataTable.Rows[$idx][$h]
            $dataTable.Rows[$idx][$h] = $dataTable.Rows[($idx - 1)][$h]
            $dataTable.Rows[($idx - 1)][$h] = $temp
        }
        $dgv.CurrentCell = $dgv.Rows[($idx - 1)].Cells[$dgv.CurrentCell.ColumnIndex]
    }

    # --- Row operation: Move Down ---
    $script:DoMoveDown = {
        $null = $dgv.EndEdit()
        $idx = & $script:GetCurrentRowIndex
        if ($idx -lt 0 -or $idx -ge ($dataTable.Rows.Count - 1)) { return }
        foreach ($h in $headers) {
            $temp = $dataTable.Rows[$idx][$h]
            $dataTable.Rows[$idx][$h] = $dataTable.Rows[($idx + 1)][$h]
            $dataTable.Rows[($idx + 1)][$h] = $temp
        }
        $dgv.CurrentCell = $dgv.Rows[($idx + 1)].Cells[$dgv.CurrentCell.ColumnIndex]
    }

    # --- Toolbar buttons ---
    $toolBtnFont = New-Object System.Drawing.Font("Consolas", 9)
    $toolBtnColor = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $toolBtnBorder = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $toolBtnY = 6
    $toolBtnH = 32

    # + Add
    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Location = New-Object System.Drawing.Point(10, $toolBtnY)
    $btnAdd.Size = New-Object System.Drawing.Size(90, $toolBtnH)
    $btnAdd.Text = "+ Add"
    $btnAdd.Font = $toolBtnFont
    $btnAdd.ForeColor = $successGreen
    $btnAdd.BackColor = $toolBtnColor
    $btnAdd.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnAdd.FlatAppearance.BorderColor = $toolBtnBorder
    $btnAdd.FlatAppearance.BorderSize = 1
    $btnAdd.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $toolPanel.Controls.Add($btnAdd)
    $btnAdd.Add_Click({ & $script:DoAddRow })

    # Copy
    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Location = New-Object System.Drawing.Point(108, $toolBtnY)
    $btnCopy.Size = New-Object System.Drawing.Size(90, $toolBtnH)
    $btnCopy.Text = "Copy"
    $btnCopy.Font = $toolBtnFont
    $btnCopy.ForeColor = $accentCyan
    $btnCopy.BackColor = $toolBtnColor
    $btnCopy.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCopy.FlatAppearance.BorderColor = $toolBtnBorder
    $btnCopy.FlatAppearance.BorderSize = 1
    $btnCopy.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $toolPanel.Controls.Add($btnCopy)
    $btnCopy.Add_Click({ & $script:DoCopyRow })

    # Delete
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Location = New-Object System.Drawing.Point(206, $toolBtnY)
    $btnDelete.Size = New-Object System.Drawing.Size(90, $toolBtnH)
    $btnDelete.Text = "Delete"
    $btnDelete.Font = $toolBtnFont
    $btnDelete.ForeColor = $errorRed
    $btnDelete.BackColor = $toolBtnColor
    $btnDelete.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDelete.FlatAppearance.BorderColor = $toolBtnBorder
    $btnDelete.FlatAppearance.BorderSize = 1
    $btnDelete.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $toolPanel.Controls.Add($btnDelete)
    $btnDelete.Add_Click({ & $script:DoDeleteRow })

    # Separator
    $sep1 = New-Object System.Windows.Forms.Label
    $sep1.Location = New-Object System.Drawing.Point(308, ($toolBtnY + 2))
    $sep1.Size = New-Object System.Drawing.Size(1, ($toolBtnH - 4))
    $sep1.BackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $null = $toolPanel.Controls.Add($sep1)

    # Up
    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Location = New-Object System.Drawing.Point(320, $toolBtnY)
    $btnUp.Size = New-Object System.Drawing.Size(70, $toolBtnH)
    $btnUp.Text = [char]0x25B2 + " Up"
    $btnUp.Font = $toolBtnFont
    $btnUp.ForeColor = $textWhite
    $btnUp.BackColor = $toolBtnColor
    $btnUp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnUp.FlatAppearance.BorderColor = $toolBtnBorder
    $btnUp.FlatAppearance.BorderSize = 1
    $btnUp.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $toolPanel.Controls.Add($btnUp)
    $btnUp.Add_Click({ & $script:DoMoveUp })

    # Down
    $btnDown = New-Object System.Windows.Forms.Button
    $btnDown.Location = New-Object System.Drawing.Point(398, $toolBtnY)
    $btnDown.Size = New-Object System.Drawing.Size(75, $toolBtnH)
    $btnDown.Text = [char]0x25BC + " Down"
    $btnDown.Font = $toolBtnFont
    $btnDown.ForeColor = $textWhite
    $btnDown.BackColor = $toolBtnColor
    $btnDown.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDown.FlatAppearance.BorderColor = $toolBtnBorder
    $btnDown.FlatAppearance.BorderSize = 1
    $btnDown.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $toolPanel.Controls.Add($btnDown)
    $btnDown.Add_Click({ & $script:DoMoveDown })

    # --- ToolTips ---
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $toolTip.ForeColor = $textWhite
    $null = $toolTip.SetToolTip($btnAdd, "Add empty row after selection (Ctrl+N)")
    $null = $toolTip.SetToolTip($btnCopy, "Duplicate selected row (Ctrl+D)")
    $null = $toolTip.SetToolTip($btnDelete, "Delete selected rows (Delete)")
    $null = $toolTip.SetToolTip($btnUp, "Move row up (Ctrl+Up)")
    $null = $toolTip.SetToolTip($btnDown, "Move row down (Ctrl+Down)")

    # --- Search controls (right side of toolbar) ---
    $anchorTR = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $tpW = $toolPanel.Width

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Search:"
    $lblSearch.Location = New-Object System.Drawing.Point(($tpW - 430), 12)
    $lblSearch.Size = New-Object System.Drawing.Size(55, 20)
    $lblSearch.Font = $toolBtnFont
    $lblSearch.ForeColor = $textGray
    $lblSearch.Anchor = $anchorTR
    $null = $toolPanel.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(($tpW - 370), 9)
    $txtSearch.Size = New-Object System.Drawing.Size(175, 26)
    $txtSearch.Font = $fontNormal
    $txtSearch.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $txtSearch.ForeColor = $textWhite
    $txtSearch.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtSearch.Anchor = $anchorTR
    $null = $toolPanel.Controls.Add($txtSearch)
    $null = $toolTip.SetToolTip($txtSearch, "Search all cells (Ctrl+F to focus, Enter to find next)")

    $btnFind = New-Object System.Windows.Forms.Button
    $btnFind.Location = New-Object System.Drawing.Point(($tpW - 188), $toolBtnY)
    $btnFind.Size = New-Object System.Drawing.Size(85, $toolBtnH)
    $btnFind.Text = "Find Next"
    $btnFind.Font = $toolBtnFont
    $btnFind.ForeColor = $warnYellow
    $btnFind.BackColor = $toolBtnColor
    $btnFind.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnFind.FlatAppearance.BorderColor = $toolBtnBorder
    $btnFind.FlatAppearance.BorderSize = 1
    $btnFind.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnFind.Anchor = $anchorTR
    $null = $toolPanel.Controls.Add($btnFind)
    $null = $toolTip.SetToolTip($btnFind, "Find next match (Enter in search box)")

    $lblMatchCount = New-Object System.Windows.Forms.Label
    $lblMatchCount.Location = New-Object System.Drawing.Point(($tpW - 98), 12)
    $lblMatchCount.Size = New-Object System.Drawing.Size(90, 20)
    $lblMatchCount.Font = $toolBtnFont
    $lblMatchCount.ForeColor = $textGray
    $lblMatchCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblMatchCount.Text = ""
    $lblMatchCount.Anchor = $anchorTR
    $null = $toolPanel.Controls.Add($lblMatchCount)

    # --- Search logic ---
    $script:SearchMatches = @()
    $script:SearchMatchIndex = -1

    $script:DoSearch = {
        $searchText = $txtSearch.Text
        $script:SearchMatches = @()
        $script:SearchMatchIndex = -1

        if ([string]::IsNullOrWhiteSpace($searchText)) {
            $lblMatchCount.Text = ""
            $lblMatchCount.ForeColor = $textGray
            return
        }

        for ($r = 0; $r -lt $dataTable.Rows.Count; $r++) {
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $cellVal = $dataTable.Rows[$r][$headers[$c]]
                if ($null -ne $cellVal -and $cellVal.ToString().IndexOf($searchText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $script:SearchMatches += ,@($r, $c)
                }
            }
        }

        if ($script:SearchMatches.Count -gt 0) {
            $script:SearchMatchIndex = 0
            $m = $script:SearchMatches[0]
            $dgv.CurrentCell = $dgv.Rows[$m[0]].Cells[$m[1]]
            $lblMatchCount.Text = "1/$($script:SearchMatches.Count)"
            $lblMatchCount.ForeColor = $successGreen
        } else {
            $lblMatchCount.Text = "0 found"
            $lblMatchCount.ForeColor = $errorRed
        }
    }

    $script:DoFindNext = {
        if ($script:SearchMatches.Count -eq 0) {
            & $script:DoSearch
            return
        }
        $script:SearchMatchIndex = ($script:SearchMatchIndex + 1) % $script:SearchMatches.Count
        $m = $script:SearchMatches[$script:SearchMatchIndex]
        $dgv.CurrentCell = $dgv.Rows[$m[0]].Cells[$m[1]]
        $lblMatchCount.Text = "$($script:SearchMatchIndex + 1)/$($script:SearchMatches.Count)"
        $lblMatchCount.ForeColor = $successGreen
    }

    $txtSearch.Add_TextChanged({ & $script:DoSearch })

    $txtSearch.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            & $script:DoFindNext
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $txtSearch.Text = ""
            $dgv.Focus()
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
    })

    $btnFind.Add_Click({ & $script:DoFindNext })

    # --- DataGridView (minimal — system default colors for reliable rendering) ---
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(0, 45)
    $dgv.Size = New-Object System.Drawing.Size(
        $formEditor.ClientSize.Width,
        ($formEditor.ClientSize.Height - 45 - 80)
    )
    $dgv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.AllowUserToOrderColumns = $false
    $dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $dgv.MultiSelect = $true
    $dgv.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnKeystrokeOrF2
    $dgv.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
    $dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $dgv.RowHeadersWidth = 50
    $dgv.ForeColor = [System.Drawing.Color]::Black

    # Bind DataTable
    $dgv.DataSource = $dataTable

    # Disable sorting on all columns
    foreach ($col in $dgv.Columns) {
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    }

    $null = $formEditor.Controls.Add($dgv)

    # --- Cell Validation ---
    $dgv.Add_CellValueChanged({
        param($sender, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }

        $colName = $headers[$e.ColumnIndex]
        $cellValue = $dgv.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Value
        $isValid = $true

        if (-not [string]::IsNullOrEmpty($cellValue)) {
            if ($colName -match '(IP|Subnet|Gateway)$' -or $colName -match '^DNS\d*$') {
                if ($cellValue -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                    $isValid = $false
                }
            }
            elseif ($colName -eq 'Enabled') {
                if ($cellValue -ne '0' -and $cellValue -ne '1') { $isValid = $false }
            }
            elseif ($colName -eq 'Order') {
                $n = 0
                if (-not [int]::TryParse($cellValue, [ref]$n)) { $isValid = $false }
            }
        }

        if (-not $isValid) {
            $dgv.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Style.BackColor = [System.Drawing.Color]::MistyRose
            $dgv.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Style.ForeColor = [System.Drawing.Color]::Red
        } else {
            # Reset to default
            $dgv.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Style.BackColor = [System.Drawing.Color]::Empty
            $dgv.Rows[$e.RowIndex].Cells[$e.ColumnIndex].Style.ForeColor = [System.Drawing.Color]::Empty
        }
    })

    # --- Status Label ---
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(10, ($formEditor.ClientSize.Height - 72))
    $lblStatus.Size = New-Object System.Drawing.Size(($formEditor.ClientSize.Width - 20), 20)
    $lblStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $lblStatus.ForeColor = $textGray
    $lblStatus.Font = $fontSmall
    $null = $formEditor.Controls.Add($lblStatus)

    # Status update function
    $script:UpdateStatus = {
        $rowCount = $dataTable.Rows.Count
        $modText = if ($script:IsModified) { " | MODIFIED" } else { "" }
        $lblStatus.Text = "$rowCount rows x $($headers.Count) columns | $($script:CsvEncoding)$modText"
        $lblStatus.ForeColor = if ($script:IsModified) { $warnYellow } else { $textGray }
    }

    & $script:UpdateStatus

    # --- Modification tracking ---
    $dataTable.Add_RowChanged({
        $script:IsModified = $true
        & $script:UpdateStatus
    })
    $dataTable.Add_RowDeleted({
        $script:IsModified = $true
        & $script:UpdateStatus
    })
    $dataTable.Add_TableNewRow({
        $script:IsModified = $true
        & $script:UpdateStatus
    })

    # --- Save function ---
    $script:SaveCsv = {
        param([string]$TargetPath, [bool]$CreateBackup)

        $null = $dgv.EndEdit()

        # Backup
        if ($CreateBackup -and (Test-Path $TargetPath)) {
            $backupPath = "$TargetPath.bak"
            try {
                Copy-Item -Path $TargetPath -Destination $backupPath -Force
            }
            catch {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "Backup creation failed: $($_.Exception.Message)`nContinue saving?",
                    "Backup Warning",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($result -eq [System.Windows.Forms.DialogResult]::No) { return $false }
            }
        }

        try {
            $outputData = @()
            foreach ($dtRow in $dataTable.Rows) {
                if ($dtRow.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
                $obj = [ordered]@{}
                foreach ($h in $headers) {
                    $obj[$h] = $dtRow[$h]
                }
                $outputData += [PSCustomObject]$obj
            }

            if ($script:CsvEncoding -eq "UTF8") {
                $outputData | Export-Csv -Path $TargetPath -NoTypeInformation -Encoding UTF8
            } else {
                $outputData | Export-Csv -Path $TargetPath -NoTypeInformation -Encoding Default
            }

            $script:IsModified = $false
            & $script:UpdateStatus
            return $true
        }
        catch {
            $null = [System.Windows.Forms.MessageBox]::Show(
                "Failed to save: $($_.Exception.Message)",
                "Save Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
    }

    # --- Bottom buttons ---
    # Save button
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(10, ($formEditor.ClientSize.Height - 48))
    $btnSave.Size = New-Object System.Drawing.Size(160, 38)
    $btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $btnSave.Text = "Save (Ctrl+S)"
    $btnSave.Font = $fontBold
    $btnSave.ForeColor = $successGreen
    $btnSave.BackColor = $btnGreenBg
    $btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSave.FlatAppearance.BorderColor = $successGreen
    $btnSave.FlatAppearance.BorderSize = 1
    $btnSave.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $formEditor.Controls.Add($btnSave)

    $btnSave.Add_Click({
        $null = & $script:SaveCsv $CsvPath $true
    })

    # Save As button
    $btnSaveAs = New-Object System.Windows.Forms.Button
    $btnSaveAs.Location = New-Object System.Drawing.Point(180, ($formEditor.ClientSize.Height - 48))
    $btnSaveAs.Size = New-Object System.Drawing.Size(130, 38)
    $btnSaveAs.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $btnSaveAs.Text = "Save As..."
    $btnSaveAs.Font = $fontSmall
    $btnSaveAs.ForeColor = $accentCyan
    $btnSaveAs.BackColor = $btnCyanBg
    $btnSaveAs.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSaveAs.FlatAppearance.BorderColor = $accentCyan
    $btnSaveAs.FlatAppearance.BorderSize = 1
    $btnSaveAs.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $formEditor.Controls.Add($btnSaveAs)

    $btnSaveAs.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $saveDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName((Resolve-Path $CsvPath).Path)
        $saveDialog.FileName = [System.IO.Path]::GetFileName($CsvPath)
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $null = & $script:SaveCsv $saveDialog.FileName $false
        }
    })

    # Cancel button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(($formEditor.ClientSize.Width - 140), ($formEditor.ClientSize.Height - 48))
    $btnCancel.Size = New-Object System.Drawing.Size(130, 38)
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Text = "Cancel"
    $btnCancel.Font = $fontSmall
    $btnCancel.ForeColor = $errorRed
    $btnCancel.BackColor = $btnRedBg
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderColor = $errorRed
    $btnCancel.FlatAppearance.BorderSize = 1
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $null = $formEditor.Controls.Add($btnCancel)

    $btnCancel.Add_Click({
        $formEditor.Close()
    })

    # --- Keyboard shortcuts ---
    $formEditor.Add_KeyDown({
        param($sender, $e)
        # Ctrl+S = Save
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::S) {
            $null = & $script:SaveCsv $CsvPath $true
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Ctrl+N = Add row
        elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::N) {
            & $script:DoAddRow
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Ctrl+D = Copy row
        elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::D) {
            & $script:DoCopyRow
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Ctrl+Up = Move row up
        elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Up) {
            & $script:DoMoveUp
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Ctrl+Down = Move row down
        elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Down) {
            & $script:DoMoveDown
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Ctrl+F = Focus search box
        elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $txtSearch.Focus()
            $txtSearch.SelectAll()
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Ctrl+V = Paste from clipboard (tab-separated, e.g. from Excel)
        elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V -and -not $dgv.IsCurrentCellInEditMode) {
            $clipText = [System.Windows.Forms.Clipboard]::GetText()
            if (-not [string]::IsNullOrEmpty($clipText)) {
                $null = $dgv.EndEdit()
                $lines = $clipText -split "`r?`n"
                $startRow = if ($dgv.CurrentCell) { $dgv.CurrentCell.RowIndex } else { 0 }
                $startCol = if ($dgv.CurrentCell) { $dgv.CurrentCell.ColumnIndex } else { 0 }

                for ($li = 0; $li -lt $lines.Count; $li++) {
                    if ([string]::IsNullOrWhiteSpace($lines[$li])) { continue }
                    $values = $lines[$li] -split "`t"
                    $targetRow = $startRow + $li

                    # Add rows if needed
                    while ($targetRow -ge $dataTable.Rows.Count) {
                        $newRow = $dataTable.NewRow()
                        foreach ($h in $headers) { $newRow[$h] = "" }
                        $null = $dataTable.Rows.Add($newRow)
                    }

                    for ($vi = 0; $vi -lt $values.Count; $vi++) {
                        $targetCol = $startCol + $vi
                        if ($targetCol -ge $headers.Count) { break }
                        $dataTable.Rows[$targetRow][$headers[$targetCol]] = $values[$vi]
                    }
                }
            }
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Delete = Delete row (only when not editing a cell)
        elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete -and -not $dgv.IsCurrentCellInEditMode) {
            & $script:DoDeleteRow
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
        # Escape = Close (only when search box is not focused)
        elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape -and -not $txtSearch.Focused) {
            $formEditor.Close()
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
    })

    # --- Unsaved changes warning ---
    $formEditor.Add_FormClosing({
        param($sender, $e)
        if ($script:IsModified) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "You have unsaved changes. Save before closing?",
                "Unsaved Changes",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            switch ($result) {
                ([System.Windows.Forms.DialogResult]::Yes) {
                    $saved = & $script:SaveCsv $CsvPath $true
                    if (-not $saved) { $e.Cancel = $true }
                }
                ([System.Windows.Forms.DialogResult]::Cancel) {
                    $e.Cancel = $true
                }
            }
        }
    })

    # Timer-based delayed refresh to ensure DataGridView renders correctly
    $refreshTimer = New-Object System.Windows.Forms.Timer
    $refreshTimer.Interval = 200
    $refreshTimer.Add_Tick({
        $refreshTimer.Stop()
        $refreshTimer.Dispose()
        $dgv.Refresh()
    })
    $formEditor.Add_Shown({
        $refreshTimer.Start()
    })

    # Show editor (modal)
    $null = $formEditor.ShowDialog()
    $formEditor.Dispose()
}

# ========================================
# Entry Point
# ========================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Fabriq CSV Editor" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Show-CsvSelector

Write-Host ""
Write-Host "[INFO] CSV Editor closed" -ForegroundColor Cyan
Write-Host ""
