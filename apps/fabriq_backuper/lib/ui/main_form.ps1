# ============================================================
# FabriqBackUper - Main Form
# Single Form with swappable Panel-based views (Mode Select /
# Backup / Restore / Progress). View modules expose
# New-<Name>View functions returning a Panel; this file owns
# the form, view switching, and shared state.
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
# Shared state (script-scope, read/write by views)
# ============================================================
$script:MainForm     = $null
$script:Hostlist     = @()
$script:CurrentHost  = $null
$script:SectionList  = @()
$script:Views        = @{}   # name -> Panel
$script:ContentArea  = $null # parent Panel that holds the active view
$script:HostLabel    = $null # header host indicator

function Switch-View {
    param([string]$Name)
    foreach ($k in $script:Views.Keys) {
        $script:Views[$k].Visible = ($k -eq $Name)
    }
    # Per-view on-show hook
    $onShowName = "Show-${Name}View"
    if (Get-Command $onShowName -ErrorAction SilentlyContinue) {
        & $onShowName
    }
}

function Update-HostHeader {
    if ($null -ne $script:HostLabel) {
        if ($null -eq $script:CurrentHost) {
            $script:HostLabel.Text = "Host: (not selected)"
        } else {
            $newName = if ($script:CurrentHost.PSObject.Properties.Name -contains 'NewPCname') {
                $script:CurrentHost.NewPCname
            } else { '' }
            $suffix = if (-not [string]::IsNullOrWhiteSpace($newName)) { " -> $newName" } else { '' }
            $script:HostLabel.Text = "Host: $($script:CurrentHost.OldPCname)$suffix"
        }
    }
}

function Start-FabriqBackuperGui {
    param(
        [Parameter(Mandatory = $true)][string]$BackuperVersion,
        [Parameter(Mandatory = $true)][string]$BackuperRoot,
        [Parameter(Mandatory = $true)][string]$FabriqRoot
    )

    $script:BackuperVersion = $BackuperVersion
    $script:BackuperRoot    = $BackuperRoot
    $script:FabriqRoot      = $FabriqRoot

    # Load hostlist + section registry up-front
    $script:Hostlist = @(Get-FabriqHostlist -FabriqRoot $FabriqRoot)
    if ($null -eq $script:Hostlist -or $script:Hostlist.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load hostlist or it is empty. Please check $FabriqRoot\kernel\csv\hostlist.csv",
            "Fabriq BackUper - Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }
    $script:SectionList = @(Get-RegisteredSections -BackuperRoot $BackuperRoot)

    # Build main form (taller in Phase 2.4 to fit destination root + extra fields)
    $form = New-Object System.Windows.Forms.Form
    Set-FormStyle -Form $form -Title "Fabriq BackUper v$BackuperVersion" -Width 760 -Height 660
    $script:MainForm = $form

    # Header bar (dark stripe with title + host indicator)
    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = "Top"
    $header.Height = 44
    $header.BackColor = $script:bgPanel
    $form.Controls.Add($header)

    $titleLbl = New-StyledLabel -Text "Fabriq BackUper" `
        -X 16 -Y 10 -Width 240 -Height 22 `
        -FgColor $script:fgWhite -Font $script:fontLarge
    $header.Controls.Add($titleLbl)

    $script:HostLabel = New-StyledLabel -Text "Host: (not selected)" `
        -X 270 -Y 14 -Width 420 -Height 18 `
        -FgColor $script:fgWhite -Font $script:fontNormal
    $header.Controls.Add($script:HostLabel)

    # Content area (fills below header)
    $content = New-Object System.Windows.Forms.Panel
    $content.Dock = "Fill"
    $content.BackColor = $script:bgForm
    $form.Controls.Add($content)
    $content.BringToFront()
    $script:ContentArea = $content

    # Build views (each returns a Panel)
    $script:Views['ModeSelect'] = New-ModeSelectView
    $script:Views['Backup']     = New-BackupView
    $script:Views['Restore']    = New-RestoreView
    $script:Views['Progress']   = New-ProgressView

    foreach ($k in $script:Views.Keys) {
        $script:Views[$k].Dock = "Fill"
        $script:Views[$k].Visible = $false
        $content.Controls.Add($script:Views[$k])
    }

    Switch-View 'ModeSelect'

    [System.Windows.Forms.Application]::Run($form)
}
