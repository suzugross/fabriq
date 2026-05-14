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
        -X 32 -Y 32 -Width 280 -Height 20 -Font $script:fontBold
    $panel.Controls.Add($hostLbl)

    $combo = New-StyledComboBox -X 32 -Y 56 -Width 540 -Height 24
    $script:ModeHostCombo = $combo
    $panel.Controls.Add($combo)

    $hostHint = New-StyledLabel -Text "(populated from kernel/csv/hostlist.csv)" `
        -X 32 -Y 84 -Width 540 -Height 18 -FgColor $script:fgDim
    $panel.Controls.Add($hostHint)

    # Mode buttons row
    $btnBackup = New-StyledButton -Text "Backup" `
        -X 96 -Y 220 -Width 200 -Height 56 -BgColor $script:bgAccent
    $btnBackup.Font = $script:fontLarge
    $panel.Controls.Add($btnBackup)

    $btnRestore = New-StyledButton -Text "Restore" `
        -X 360 -Y 220 -Width 200 -Height 56 -BgColor $script:bgAdd
    $btnRestore.ForeColor = $script:fgWhite
    $btnRestore.Font = $script:fontLarge
    $panel.Controls.Add($btnRestore)

    # Footer (Quit)
    $btnQuit = New-StyledButton -Text "Quit" `
        -X 472 -Y 420 -Width 100 -Height 30
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
