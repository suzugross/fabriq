# ============================================================
# FabriqBackUper - Mode Select View
# Initial landing screen: pick a host from hostlist, then
# choose Backup or Restore.
# ============================================================

$script:ModeHostCombo  = $null

function New-ModeSelectView {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.BackColor = $script:bgForm

    # Host selector group
    $hostLbl = New-StyledLabel -Text "Select host (OldPCname):" `
        -X 40 -Y 60 -Width 280 -Height 22 -Font $script:fontBold
    $panel.Controls.Add($hostLbl)

    $combo = New-StyledComboBox -X 40 -Y 88 -Width 800 -Height 24
    $script:ModeHostCombo = $combo
    $panel.Controls.Add($combo)

    $hostHint = New-StyledLabel -Text "(populated from kernel/csv/hostlist.csv via Import-ModuleCsv + master passphrase)" `
        -X 40 -Y 118 -Width 800 -Height 18 -FgColor $script:fgDim
    $panel.Controls.Add($hostHint)

    # Mode buttons row (centered, large)
    $btnBackup = New-StyledButton -Text "Backup" `
        -X 200 -Y 240 -Width 240 -Height 72 -BgColor $script:bgAccent
    $btnBackup.Font = $script:fontTitle
    $panel.Controls.Add($btnBackup)

    $btnRestore = New-StyledButton -Text "Restore" `
        -X 500 -Y 240 -Width 240 -Height 72 -BgColor $script:bgAdd
    $btnRestore.ForeColor = $script:fgWhite
    $btnRestore.Font = $script:fontTitle
    $panel.Controls.Add($btnRestore)

    # Footer (Quit)
    $btnQuit = New-StyledButton -Text "Quit" `
        -X 740 -Y 624 -Width 100 -Height 32
    $panel.Controls.Add($btnQuit)

    # Events
    $btnBackup.Add_Click({
        if (-not (Confirm-ModeSelectHost)) { return }
        Switch-View 'Backup'
    })
    $btnRestore.Add_Click({
        if (-not (Confirm-ModeSelectHost)) { return }
        Switch-View 'Restore'
    })
    $btnQuit.Add_Click({
        $script:MainForm.Close()
    })

    return $panel
}

function Confirm-ModeSelectHost {
    if ($null -eq $script:ModeHostCombo -or $script:ModeHostCombo.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a host first.",
            "Fabriq BackUper",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $false
    }
    $script:CurrentHost = $script:Hostlist[$script:ModeHostCombo.SelectedIndex]
    Update-HostHeader
    return $true
}

function Show-ModeSelectView {
    # Populate / refresh host combo on entry
    if ($null -eq $script:ModeHostCombo) { return }
    $combo = $script:ModeHostCombo
    $combo.BeginUpdate()
    $combo.Items.Clear()
    foreach ($h in $script:Hostlist) {
        $label = $h.OldPCname
        if ($h.PSObject.Properties.Name -contains 'NewPCname' -and `
            -not [string]::IsNullOrWhiteSpace($h.NewPCname)) {
            $label = "$($h.OldPCname)   ->   $($h.NewPCname)"
        }
        [void]$combo.Items.Add($label)
    }
    $combo.EndUpdate()

    # Inherit previously selected host if any
    if ($null -ne $script:CurrentHost) {
        for ($i = 0; $i -lt $script:Hostlist.Count; $i++) {
            if ($script:Hostlist[$i].OldPCname -eq $script:CurrentHost.OldPCname) {
                $combo.SelectedIndex = $i
                break
            }
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($env:SELECTED_OLD_PCNAME)) {
        # Inherit env (parent fabriq session)
        for ($i = 0; $i -lt $script:Hostlist.Count; $i++) {
            if ($script:Hostlist[$i].OldPCname -eq $env:SELECTED_OLD_PCNAME) {
                $combo.SelectedIndex = $i
                $script:CurrentHost = $script:Hostlist[$i]
                Update-HostHeader
                break
            }
        }
    }
}
