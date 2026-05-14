# ============================================================
# FabriqBackUper - Backup View (Phase 2.2.1)
# Section toggles + per-printer selection grid + Start.
# Per-printer checkboxes are ephemeral (this run only); the
# printer section receives selected names via SectionParams.
# ============================================================

$script:BackupSectionChecks   = @{}
$script:BackupPrinterGrid     = $null
$script:BackupPrinterRows     = @()
$script:BackupEntryGrid       = $null
$script:BackupSectionContainer = $null
$script:BackupDestinationBox   = $null

# Virtual printer detection (unchecked by default)
$script:VirtualDriverPatterns = @(
    'Microsoft Print To PDF',
    'Microsoft XPS Document Writer',
    'Microsoft Shared Fax Driver',
    'Microsoft OpenXPS Class Driver',
    'OneNote',
    'Remote Desktop Easy Print'
)
$script:VirtualPortPatterns = @(
    'PORTPROMPT:', 'XPSPort:', 'FAX:', 'nul:', 'SHRFAX:'
)

function Test-BackupViewVirtualPrinter {
    param($P)
    foreach ($pat in $script:VirtualDriverPatterns) { if ($P.DriverName -like "*$pat*") { return $true } }
    foreach ($pat in $script:VirtualPortPatterns)   { if ($P.PortName   -like "*$pat*") { return $true } }
    if ($P.PortName -like 'OneNote*') { return $true }
    if ($P.PortName -match '^TS\d+$') { return $true }
    return $false
}

function New-BackupView {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.BackColor = $script:bgForm

    $btnBack = New-StyledButton -Text "< Back" -X 16 -Y 12 -Width 80 -Height 28
    $btnBack.Add_Click({ Switch-View 'ModeSelect' })
    $panel.Controls.Add($btnBack)

    $title = New-StyledLabel -Text "Backup" -X 110 -Y 14 -Width 200 -Height 24 -Font $script:fontLarge
    $panel.Controls.Add($title)

    # Destination root (Phase 2.4)
    $destLbl = New-StyledLabel -Text "Destination root (UNC OK):" `
        -X 32 -Y 50 -Width 240 -Height 18 -Font $script:fontBold -FgColor $script:fgHeader
    $panel.Controls.Add($destLbl)

    $destBox = New-Object System.Windows.Forms.TextBox
    $destBox.Location = New-Object System.Drawing.Point(32, 70)
    $destBox.Size = New-Object System.Drawing.Size(560, 24)
    Set-TextBoxStyle -TextBox $destBox
    $destBox.Text = (Join-Path $script:BackuperRoot 'Backup')
    $panel.Controls.Add($destBox)
    $script:BackupDestinationBox = $destBox

    $btnDestBrowse = New-StyledButton -Text "Browse..." -X 600 -Y 68 -Width 80 -Height 26
    $btnDestBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select backup destination root (local or UNC)"
        $dlg.ShowNewFolderButton = $true
        if (-not [string]::IsNullOrWhiteSpace($script:BackupDestinationBox.Text) -and `
            (Test-Path $script:BackupDestinationBox.Text)) {
            $dlg.SelectedPath = $script:BackupDestinationBox.Text
        }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:BackupDestinationBox.Text = $dlg.SelectedPath
        }
    })
    $panel.Controls.Add($btnDestBrowse)

    # Sections row
    $sectionGroupLbl = New-StyledLabel -Text "Sections" -X 32 -Y 108 -Width 200 -Height 18 -Font $script:fontBold -FgColor $script:fgHeader
    $panel.Controls.Add($sectionGroupLbl)
    $script:BackupSectionContainer = New-Object System.Windows.Forms.Panel
    $script:BackupSectionContainer.Location = New-Object System.Drawing.Point(32, 130)
    $script:BackupSectionContainer.Size = New-Object System.Drawing.Size(640, 30)
    $script:BackupSectionContainer.BackColor = [System.Drawing.Color]::Transparent
    $panel.Controls.Add($script:BackupSectionContainer)

    # Printer list (when Printer section is enabled)
    $pLbl = New-StyledLabel -Text "Printers on this PC (uncheck to exclude from this backup)" `
        -X 32 -Y 170 -Width 540 -Height 18 -Font $script:fontBold -FgColor $script:fgHeader
    $panel.Controls.Add($pLbl)

    $btnSelectAll = New-StyledButton -Text "Select All" -X 498 -Y 166 -Width 86 -Height 22
    $btnSelectAll.Add_Click({ Set-AllPrinterChecks $true })
    $panel.Controls.Add($btnSelectAll)
    $btnNone = New-StyledButton -Text "None" -X 588 -Y 166 -Width 60 -Height 22
    $btnNone.Add_Click({ Set-AllPrinterChecks $false })
    $panel.Controls.Add($btnNone)
    $btnRefresh = New-StyledButton -Text "Refresh" -X 652 -Y 166 -Width 60 -Height 22
    $btnRefresh.Add_Click({ Update-BackupPrinterGrid })
    $panel.Controls.Add($btnRefresh)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(32, 194)
    $grid.Size = New-Object System.Drawing.Size(680, 180)
    Set-GridStyle -Grid $grid
    $grid.ReadOnly = $false

    $colCk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCk.HeaderText = ""
    $colCk.Width = 32
    $colCk.Name = "Check"
    [void]$grid.Columns.Add($colCk)

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Printer Name"
    $colName.Width = 280
    $colName.Name = "Name"
    $colName.ReadOnly = $true
    [void]$grid.Columns.Add($colName)

    $colDriver = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDriver.HeaderText = "Driver"
    $colDriver.Width = 180
    $colDriver.Name = "Driver"
    $colDriver.ReadOnly = $true
    [void]$grid.Columns.Add($colDriver)

    $colPort = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPort.HeaderText = "Port"
    $colPort.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colPort.Name = "Port"
    $colPort.ReadOnly = $true
    [void]$grid.Columns.Add($colPort)

    $panel.Controls.Add($grid)
    $script:BackupPrinterGrid = $grid

    # User Data entries preview (still preview-only in 2.4; editor is Phase 2.2.2)
    $entryLbl = New-StyledLabel -Text "User Data Entries (preview; editor in Phase 2.2.2)" `
        -X 32 -Y 384 -Width 480 -Height 18 -Font $script:fontBold -FgColor $script:fgHeader
    $panel.Controls.Add($entryLbl)

    $eGrid = New-Object System.Windows.Forms.DataGridView
    $eGrid.Location = New-Object System.Drawing.Point(32, 406)
    $eGrid.Size = New-Object System.Drawing.Size(680, 80)
    Set-GridStyle -Grid $eGrid

    $ec1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $ec1.HeaderText = "On"; $ec1.Width = 36; $ec1.DefaultCellStyle.Alignment = "MiddleCenter"
    [void]$eGrid.Columns.Add($ec1)
    $ec2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $ec2.HeaderText = "Description"; $ec2.Width = 200
    [void]$eGrid.Columns.Add($ec2)
    $ec3 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $ec3.HeaderText = "SourcePath"
    $ec3.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$eGrid.Columns.Add($ec3)
    $panel.Controls.Add($eGrid)
    $script:BackupEntryGrid = $eGrid

    # Start button
    $btnStart = New-StyledButton -Text "Start Backup" -X 512 -Y 496 -Width 200 -Height 40 -BgColor $script:bgAccent
    $btnStart.Font = $script:fontLarge
    $btnStart.Add_Click({ Invoke-BackupStart })
    $panel.Controls.Add($btnStart)

    return $panel
}

function Set-AllPrinterChecks {
    param([bool]$Checked)
    if ($null -eq $script:BackupPrinterGrid) { return }
    foreach ($row in $script:BackupPrinterGrid.Rows) {
        $row.Cells['Check'].Value = $Checked
    }
}

function Update-BackupPrinterGrid {
    if ($null -eq $script:BackupPrinterGrid) { return }
    $grid = $script:BackupPrinterGrid
    $grid.Rows.Clear()
    $script:BackupPrinterRows = @()
    $allPrinters = @()
    try { $allPrinters = @(Get-Printer -ErrorAction Stop) } catch { return }
    foreach ($p in $allPrinters) {
        $isVirtual = Test-BackupViewVirtualPrinter -P $p
        $defaultChecked = -not $isVirtual
        $null = $grid.Rows.Add($defaultChecked, $p.Name, $p.DriverName, $p.PortName)
        $script:BackupPrinterRows += $p
    }
}

function Show-BackupView {
    # Refresh section checkboxes
    $cont = $script:BackupSectionContainer
    $cont.Controls.Clear()
    $script:BackupSectionChecks = @{}
    $x = 0
    foreach ($s in $script:SectionList) {
        $cb = New-StyledCheckBox -Text $s.DisplayName -X $x -Y 4 -Width 280 -Height 22 -Checked ($s.Enabled -eq "1")
        $cb.Tag = $s.SectionName
        $cont.Controls.Add($cb)
        $script:BackupSectionChecks[$s.SectionName] = $cb
        $x += 300
    }

    # Populate printer grid
    Update-BackupPrinterGrid

    # Populate userdata preview from FabriqBackUper-owned CSV (Phase 2.3)
    $grid = $script:BackupEntryGrid
    $grid.Rows.Clear()
    $userdataCsv = Join-Path $script:BackuperRoot 'data\userdata_list.csv'
    if (Test-Path $userdataCsv) {
        try {
            $rows = Import-Csv -Path $userdataCsv
            foreach ($r in $rows) {
                $onMark = if ($r.Enabled -eq "1") { "[v]" } else { "[ ]" }
                [void]$grid.Rows.Add($onMark, $r.Description, $r.SourcePath)
            }
        } catch { [void]$grid.Rows.Add("!", "Failed to read CSV", $_.Exception.Message) }
    } else {
        [void]$grid.Rows.Add("?", "(not found)", $userdataCsv)
    }
}

function Invoke-BackupStart {
    if ($null -eq $script:CurrentHost) {
        [System.Windows.Forms.MessageBox]::Show("No host selected.", "Fabriq BackUper",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $picked = @()
    foreach ($s in $script:SectionList) {
        $cb = $script:BackupSectionChecks[$s.SectionName]
        if ($null -ne $cb -and $cb.Checked) { $picked += $s }
    }
    if ($picked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No sections selected.", "Fabriq BackUper",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # Collect printer selection
    $selectedPrinters = @()
    if ($null -ne $script:BackupPrinterGrid) {
        foreach ($row in $script:BackupPrinterGrid.Rows) {
            if ($row.Cells['Check'].Value -eq $true) {
                $selectedPrinters += [string]$row.Cells['Name'].Value
            }
        }
    }

    $sectionParams = @{
        printer = @{
            IncludePrinters       = $selectedPrinters
            IncludeDriverBinaries = $true
            IncludePrintSettings  = $true
        }
    }

    # Phase 2.4: destination root from UI
    $destRoot = $script:BackupDestinationBox.Text
    if ([string]::IsNullOrWhiteSpace($destRoot)) {
        $destRoot = Join-Path $script:BackuperRoot 'Backup'
    }
    # If UNC, ensure reachability (prompt for credentials if needed)
    if (-not (Resolve-UncAccess -Path $destRoot)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Cannot reach destination: $destRoot",
            "Fabriq BackUper", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $printerSummary = if ($selectedPrinters.Count -gt 0) {
        "Printers: $($selectedPrinters.Count) selected"
    } else {
        "Printers: 0 (printer section will be skipped)"
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Start backup for $($script:CurrentHost.OldPCname)?`n`nDestination: $destRoot`nSections: $(@($picked | ForEach-Object { $_.SectionName }) -join ', ')`n$printerSummary",
        "Fabriq BackUper - Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    Switch-View 'Progress'
    Initialize-ProgressView -Title "Backup in progress..."
    Append-ProgressLog "Starting backup for $($script:CurrentHost.OldPCname)"
    Append-ProgressLog "Destination: $destRoot"
    if ($selectedPrinters.Count -gt 0) {
        Append-ProgressLog "Selected printers: $($selectedPrinters -join ', ')"
    }
    $script:MainForm.Refresh()

    $result = Invoke-BackuperBackupCore `
        -SelectedHost $script:CurrentHost `
        -PickedSections $picked `
        -BackuperRoot $script:BackuperRoot `
        -FabriqRoot $script:FabriqRoot `
        -BackuperVersion $script:BackuperVersion `
        -SectionParamsBySection $sectionParams `
        -DestinationRoot $destRoot

    Append-ProgressLog ""
    Append-ProgressLog "=========================================="
    Append-ProgressLog "Backup complete: $($result.Status)"
    Append-ProgressLog "$($result.Message)"
    foreach ($key in $result.SectionResults.Keys) {
        $r = $result.SectionResults[$key]
        Append-ProgressLog ("  [{0,-10}] {1,-8} ({2} ms)" -f $key, $r.Status, $r.ElapsedMs)
        if ($r.InternalSectionDir) {
            Append-ProgressLog "             -> $($r.InternalSectionDir)"
        } elseif ($r.ExternalOutputDir) {
            Append-ProgressLog "             -> $($r.ExternalOutputDir)  (external)"
        }
    }
    Set-ProgressFinished
}
