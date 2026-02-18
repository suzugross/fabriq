# ========================================
# Report Viewer
# Displays a generated HTML checklist in a
# Windows Forms WebBrowser window.
# Usage: .\view_report.ps1 -HtmlPath "C:\...\checklist_xxx.html"
# ========================================
param(
    [string]$HtmlPath
)

Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing       -ErrorAction SilentlyContinue

# ========================================
# File existence check
# ========================================
if ([string]::IsNullOrEmpty($HtmlPath) -or -not (Test-Path $HtmlPath)) {
    $msg = "Report file not found: $HtmlPath"
    if (Get-Command Show-Error -ErrorAction SilentlyContinue) {
        Show-Error $msg
    } else {
        Write-Host "[ERROR] $msg" -ForegroundColor Red
    }
    return
}

$absolutePath = (Resolve-Path $HtmlPath).Path

# ========================================
# IE11 emulation (process-scoped registry)
# WebBrowser control defaults to IE7 mode;
# IE11 emulation (0x2EE1) is required for
# modern CSS (flexbox / grid) to render.
# ========================================
try {
    $emulKey = "HKCU:\SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
    $exeName = [System.IO.Path]::GetFileName(
        [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    )
    if (-not (Test-Path $emulKey)) { $null = New-Item -Path $emulKey -Force }
    Set-ItemProperty -Path $emulKey -Name $exeName -Value 0x00002EE1 -Type DWord -ErrorAction Stop
}
catch { <# non-fatal: display may fall back to IE7 mode #> }

# ========================================
# Show viewer window
# ========================================
function Show-ReportViewer {
    param([string]$FilePath)

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "fabriq Report Viewer"
    $form.Size            = New-Object System.Drawing.Size(1200, 860)
    $form.MinimumSize     = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition   = "CenterScreen"
    $form.BackColor       = [System.Drawing.Color]::FromArgb(26, 26, 46)

    $browser = New-Object System.Windows.Forms.WebBrowser
    $browser.Dock                  = [System.Windows.Forms.DockStyle]::Fill
    $browser.ScriptErrorsSuppressed = $true
    $browser.Url                   = [System.Uri]("file:///" + $FilePath.Replace('\', '/'))

    $form.Controls.Add($browser)

    # ESC to close
    $form.KeyPreview = $true
    $form.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $form.Close() }
    })

    $form.Add_Shown({ $form.Activate() })

    [void]$form.ShowDialog()
    $browser.Dispose()
    $form.Dispose()
}

Show-ReportViewer -FilePath $absolutePath
