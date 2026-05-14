# ============================================================
# FabriqBackUper - Restore View (Phase 2.2.1)
# Pick a backup timestamp + sections + per-printer selection
# (sourced from manifest of the selected backup).
# ============================================================

$script:RestoreTimestampCombo  = $null
$script:RestoreSectionChecks   = @{}
$script:RestoreManifestLabel   = $null
$script:RestoreSectionContainer = $null
$script:RestorePrinterGrid     = $null
$script:RestoreCurrentManifest = $null
$script:RestoreExplicitDir     = $null   # Phase 2.4: when Browse used
$script:RestoreBrowseLabel     = $null

function New-RestoreView {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.BackColor = $script:bgForm

    $btnBack = New-StyledButton -Text "< Back" -X 16 -Y 12 -Width 80 -Height 28
    $btnBack.Add_Click({ Switch-View 'ModeSelect' })
    $panel.Controls.Add($btnBack)

    $title = New-StyledLabel -Text "Restore" -X 110 -Y 14 -Width 200 -Height 24 -Font $script:fontLarge
    $panel.Controls.Add($title)

    # Timestamp group
    $tsLbl = New-StyledLabel -Text "Backup Timestamp" -X 32 -Y 50 -Width 240 -Height 18 -Font $script:fontBold -FgColor $script:fgHeader
    $panel.Controls.Add($tsLbl)

    $combo = New-StyledComboBox -X 32 -Y 72 -Width 360 -Height 24
    $combo.Add_SelectedIndexChanged({
        $script:RestoreExplicitDir = $null    # picking dropdown overrides any prior Browse
        if ($null -ne $script:RestoreBrowseLabel) { $script:RestoreBrowseLabel.Text = "" }
        Update-RestoreSelection
    })
    $script:RestoreTimestampCombo = $combo
    $panel.Controls.Add($combo)

    # Phase 2.4: Browse for backup folder (UNC OK)
    $btnBrowse = New-StyledButton -Text "Browse for backup..." -X 400 -Y 70 -Width 180 -Height 26
    $btnBrowse.Add_Click({ Invoke-RestoreBrowse })
    $panel.Controls.Add($btnBrowse)

    $script:RestoreBrowseLabel = New-StyledLabel -Text "" -X 32 -Y 100 -Width 700 -Height 16 -FgColor $script:fgDim
    $panel.Controls.Add($script:RestoreBrowseLabel)

    $script:RestoreManifestLabel = New-StyledLabel -Text "" -X 32 -Y 118 -Width 700 -Height 32 -FgColor $script:fgDim
    $panel.Controls.Add($script:RestoreManifestLabel)

    # Sections row
    $sectionLbl = New-StyledLabel -Text "Sections" -X 32 -Y 160 -Width 240 -Height 18 -Font $script:fontBold -FgColor $script:fgHeader
    $panel.Controls.Add($sectionLbl)
    $script:RestoreSectionContainer = New-Object System.Windows.Forms.Panel
    $script:RestoreSectionContainer.Location = New-Object System.Drawing.Point(32, 182)
    $script:RestoreSectionContainer.Size = New-Object System.Drawing.Size(680, 30)
    $script:RestoreSectionContainer.BackColor = [System.Drawing.Color]::Transparent
    $panel.Controls.Add($script:RestoreSectionContainer)

    # Printer list (from manifest of selected backup)
    $pLbl = New-StyledLabel -Text "Printers in this backup (uncheck to exclude from restore)" `
        -X 32 -Y 222 -Width 540 -Height 18 -Font $script:fontBold -FgColor $script:fgHeader
    $panel.Controls.Add($pLbl)

    $btnSelAll = New-StyledButton -Text "Select All" -X 498 -Y 218 -Width 86 -Height 22
    $btnSelAll.Add_Click({ Set-AllRestorePrinterChecks $true })
    $panel.Controls.Add($btnSelAll)
    $btnNone = New-StyledButton -Text "None" -X 588 -Y 218 -Width 60 -Height 22
    $btnNone.Add_Click({ Set-AllRestorePrinterChecks $false })
    $panel.Controls.Add($btnNone)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(32, 246)
    $grid.Size = New-Object System.Drawing.Size(680, 240)
    Set-GridStyle -Grid $grid
    $grid.ReadOnly = $false

    $colCk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCk.HeaderText = ""; $colCk.Width = 32; $colCk.Name = "Check"
    [void]$grid.Columns.Add($colCk)
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Printer Name"; $colName.Width = 260; $colName.Name = "Name"; $colName.ReadOnly = $true
    [void]$grid.Columns.Add($colName)
    $colDriver = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDriver.HeaderText = "Driver"; $colDriver.Width = 180; $colDriver.Name = "Driver"; $colDriver.ReadOnly = $true
    [void]$grid.Columns.Add($colDriver)
    $colPort = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPort.HeaderText = "Port"; $colPort.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colPort.Name = "Port"; $colPort.ReadOnly = $true
    [void]$grid.Columns.Add($colPort)
    $panel.Controls.Add($grid)
    $script:RestorePrinterGrid = $grid

    # Start button
    $btnStart = New-StyledButton -Text "Start Restore" -X 512 -Y 496 -Width 200 -Height 40 -BgColor $script:bgAdd
    $btnStart.ForeColor = $script:fgWhite
    $btnStart.Font = $script:fontLarge
    $btnStart.Add_Click({ Invoke-RestoreStart })
    $panel.Controls.Add($btnStart)

    return $panel
}

function Invoke-RestoreBrowse {
    # Phase 2.4: pick an arbitrary backup folder (local or UNC) that contains
    # a fabriq-backuper-snapshot manifest.json.
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select backup folder (must contain manifest.json)"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $chosen = $dlg.SelectedPath

    # UNC auth if needed
    if (-not (Resolve-UncAccess -Path $chosen)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Cannot reach folder: $chosen",
            "Fabriq BackUper", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # Validate manifest.json
    $mfPath = Join-Path $chosen 'manifest.json'
    if (-not (Test-Path $mfPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "manifest.json not found in:`n$chosen",
            "Fabriq BackUper", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    $agg = $null
    try { $agg = Get-Content -Path $mfPath -Raw | ConvertFrom-Json }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to parse manifest.json: $($_.Exception.Message)",
            "Fabriq BackUper", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    if ($agg.manifestType -ne 'fabriq-backuper-snapshot') {
        [System.Windows.Forms.MessageBox]::Show(
            "Unexpected manifestType: $($agg.manifestType)",
            "Fabriq BackUper", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # Accepted: switch into Browse mode
    $script:RestoreExplicitDir = $chosen
    $script:RestoreTimestampCombo.SelectedIndex = -1
    $script:RestoreBrowseLabel.Text = "Browse mode: $chosen"
    # Show aggregate summary
    $sz = if ($agg.summary.totalBytes) { [math]::Round([long]$agg.summary.totalBytes / 1MB, 1) } else { 0 }
    $secCount = if ($agg.summary.sectionCount) { [int]$agg.summary.sectionCount } else { 0 }
    $script:RestoreManifestLabel.Text = "aggregate manifest  |  collectedAt=$($agg.collectedAt)  |  oldPcName=$($agg.oldPcName)  |  sections=$secCount  |  totalBytes=$sz MB"

    # Populate printer list from internal manifest (if present)
    Show-RestorePrinterListFromAggregate -AggregateDir $chosen
}

function Show-RestorePrinterListFromAggregate {
    param([Parameter(Mandatory = $true)][string]$AggregateDir)
    $script:RestorePrinterGrid.Rows.Clear()
    $printerManifestPath = Join-Path $AggregateDir 'sections\printer\manifest.json'
    if (-not (Test-Path $printerManifestPath)) {
        $null = $script:RestorePrinterGrid.Rows.Add($false, "(no printer section manifest found)", "", "")
        return
    }
    try {
        $pm = Get-Content -Path $printerManifestPath -Raw | ConvertFrom-Json
        foreach ($p in @($pm.items.printers)) {
            if ($p.driverName -eq 'Remote Desktop Easy Print') { continue }
            if ($p.portName -match '^TS\d+$') { continue }
            $isVirtual = $false
            $virtPats = @('Microsoft Print To PDF','Microsoft XPS Document Writer','OneNote','Microsoft Shared Fax','Microsoft OpenXPS')
            foreach ($vp in $virtPats) { if ($p.driverName -like "*$vp*") { $isVirtual = $true; break } }
            $virtPortPats = @('PORTPROMPT:','XPSPort:','FAX:','nul:','SHRFAX:')
            foreach ($vp in $virtPortPats) { if ($p.portName -like "*$vp*") { $isVirtual = $true; break } }
            if ($p.portName -like 'OneNote*') { $isVirtual = $true }
            $defaultChecked = -not $isVirtual
            $null = $script:RestorePrinterGrid.Rows.Add($defaultChecked, $p.name, $p.driverName, $p.portName)
        }
    } catch { }
}

function Set-AllRestorePrinterChecks {
    param([bool]$Checked)
    if ($null -eq $script:RestorePrinterGrid) { return }
    foreach ($row in $script:RestorePrinterGrid.Rows) {
        $row.Cells['Check'].Value = $Checked
    }
}

function Show-RestoreView {
    if ($null -eq $script:CurrentHost) {
        [System.Windows.Forms.MessageBox]::Show("No host selected.", "Fabriq BackUper",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        Switch-View 'ModeSelect'
        return
    }

    # Populate timestamp combo
    $combo = $script:RestoreTimestampCombo
    $combo.Items.Clear()
    $timestamps = Get-BackupTimestamps -BackuperRoot $script:BackuperRoot -OldPcName $script:CurrentHost.OldPCname
    foreach ($ts in $timestamps) { [void]$combo.Items.Add($ts) }
    if ($timestamps.Count -gt 0) { $combo.SelectedIndex = 0 }
    else { $script:RestoreManifestLabel.Text = "(no backups found for $($script:CurrentHost.OldPCname))" }

    # Section checkboxes
    $cont = $script:RestoreSectionContainer
    $cont.Controls.Clear()
    $script:RestoreSectionChecks = @{}
    $x = 0
    foreach ($s in $script:SectionList) {
        $cb = New-StyledCheckBox -Text $s.DisplayName -X $x -Y 4 -Width 280 -Height 22 -Checked ($s.Enabled -eq "1")
        $cb.Tag = $s.SectionName
        $cont.Controls.Add($cb)
        $script:RestoreSectionChecks[$s.SectionName] = $cb
        $x += 300
    }
}

function Update-RestoreSelection {
    # Hostlist-driven path. Browse path is handled separately by
    # Invoke-RestoreBrowse / Show-RestorePrinterListFromAggregate.
    if ($script:RestorePrinterGrid) { $script:RestorePrinterGrid.Rows.Clear() }

    if ($null -eq $script:RestoreTimestampCombo -or $script:RestoreTimestampCombo.SelectedIndex -lt 0) {
        if ([string]::IsNullOrWhiteSpace($script:RestoreExplicitDir)) {
            $script:RestoreManifestLabel.Text = ""
        }
        return
    }
    $ts = $script:RestoreTimestampCombo.SelectedItem
    $aggregateDir = Join-Path (Join-Path (Join-Path $script:BackuperRoot 'Backup') $script:CurrentHost.OldPCname) $ts
    $aggregatePath = Join-Path $aggregateDir 'manifest.json'

    if (-not (Test-Path $aggregatePath)) {
        $script:RestoreManifestLabel.Text = "(manifest.json not found in $ts)"
        return
    }
    try {
        $agg = Get-Content -Path $aggregatePath -Raw | ConvertFrom-Json
        $sz = if ($agg.summary.totalBytes) { [math]::Round([long]$agg.summary.totalBytes / 1MB, 1) } else { 0 }
        $secCount = if ($agg.summary.sectionCount) { [int]$agg.summary.sectionCount } else { 0 }
        $script:RestoreManifestLabel.Text = "aggregate manifest  |  collectedAt=$($agg.collectedAt)  |  sections=$secCount  |  totalBytes=$sz MB"
    }
    catch {
        $script:RestoreManifestLabel.Text = "aggregate manifest parse failed: $($_.Exception.Message)"
    }

    Show-RestorePrinterListFromAggregate -AggregateDir $aggregateDir
}

function Invoke-RestoreStart {
    # Determine source mode: Browse explicit dir vs hostlist-driven timestamp
    $useExplicit = -not [string]::IsNullOrWhiteSpace($script:RestoreExplicitDir)
    if (-not $useExplicit) {
        if ($null -eq $script:CurrentHost) { return }
        if ($null -eq $script:RestoreTimestampCombo -or $script:RestoreTimestampCombo.SelectedIndex -lt 0) {
            [System.Windows.Forms.MessageBox]::Show("Select a backup timestamp or use 'Browse for backup...'.", "Fabriq BackUper",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
    }

    $picked = @()
    foreach ($s in $script:SectionList) {
        $cb = $script:RestoreSectionChecks[$s.SectionName]
        if ($null -ne $cb -and $cb.Checked) { $picked += $s }
    }
    if ($picked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No sections selected.", "Fabriq BackUper",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # Collect printer selection (from internal manifest if available)
    $selectedPrinters = @()
    if ($null -ne $script:RestorePrinterGrid) {
        foreach ($row in $script:RestorePrinterGrid.Rows) {
            if ($row.Cells['Check'].Value -eq $true) {
                $name = [string]$row.Cells['Name'].Value
                if (-not [string]::IsNullOrWhiteSpace($name) -and $name -notlike '(no *') {
                    $selectedPrinters += $name
                }
            }
        }
    }

    $sectionParams = @{
        printer = @{ IncludePrinters = $selectedPrinters }
    }

    # Build host context: from CurrentHost (Hostlist mode) or synthesized
    # from manifest.oldPcName (Browse mode).
    $hostForEngine = $script:CurrentHost
    $sourceLabel = ""
    if ($useExplicit) {
        $aggMfPath = Join-Path $script:RestoreExplicitDir 'manifest.json'
        try {
            $agg = Get-Content -Path $aggMfPath -Raw | ConvertFrom-Json
            $hostForEngine = [PSCustomObject]@{ OldPCname = $agg.oldPcName }
        } catch {
            $hostForEngine = [PSCustomObject]@{ OldPCname = '(unknown)' }
        }
        $sourceLabel = "Browse: $($script:RestoreExplicitDir)"
    } else {
        $ts = $script:RestoreTimestampCombo.SelectedItem
        $sourceLabel = "Hostlist: $($script:CurrentHost.OldPCname) / $ts"
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Restore from:`n  $sourceLabel`n`nSections: $(@($picked | ForEach-Object { $_.SectionName }) -join ', ')`nPrinters: $($selectedPrinters.Count) selected",
        "Fabriq BackUper - Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    Switch-View 'Progress'
    Initialize-ProgressView -Title "Restore in progress..."
    Append-ProgressLog "Restoring from: $sourceLabel"
    if ($selectedPrinters.Count -gt 0) {
        Append-ProgressLog "Selected printers: $($selectedPrinters -join ', ')"
    }
    $script:MainForm.Refresh()

    $coreArgs = @{
        SelectedHost = $hostForEngine
        PickedSections = $picked
        BackuperRoot = $script:BackuperRoot
        FabriqRoot = $script:FabriqRoot
        SectionParamsBySection = $sectionParams
    }
    if ($useExplicit) {
        $coreArgs.ExplicitAggregateDir = $script:RestoreExplicitDir
    } else {
        $coreArgs.PickedTimestamp = $script:RestoreTimestampCombo.SelectedItem
    }
    $result = Invoke-BackuperRestoreCore @coreArgs

    Append-ProgressLog ""
    Append-ProgressLog "=========================================="
    Append-ProgressLog "Restore complete: $($result.Status)"
    Append-ProgressLog "$($result.Message)"
    foreach ($key in $result.SectionResults.Keys) {
        $r = $result.SectionResults[$key]
        Append-ProgressLog ("  [{0,-10}] {1,-8} ({2} ms)" -f $key, $r.Status, $r.ElapsedMs)
    }
    Set-ProgressFinished
}
