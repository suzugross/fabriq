# ========================================
# Domain Join Script
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Show-Info "Executing domain join process..."
Write-Host ""

# ========================================
# Load domain.csv
# ========================================
$csvPath = Join-Path $PSScriptRoot "domain.csv"

$domainList = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "domain", "user", "pass", "dns")
if ($null -eq $domainList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load domain.csv")
}

$domainEntry = $domainList | Where-Object { $_.Enabled -eq '1' } | Select-Object -First 1
if ($null -eq $domainEntry) {
    Show-Info "No enabled entries in domain.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}
$DOMAIN = $domainEntry.'domain'
$USER = $domainEntry.'user'
$PASS = $domainEntry.'pass'
$DNS = $domainEntry.'dns'

# ========================================
# Helper: Show error dialog with input box
# Returns the text entered in the input box
# ========================================
function Show-ErrorDialog {
    param(
        [string]$Title,
        [string]$Message
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(520, 380)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(460, 230)
    $label.Text = $Message
    $label.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(20, 260)
    $textBox.Size = New-Object System.Drawing.Size(340, 25)
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(380, 258)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = "Retry"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $form.Add_Shown({ $form.Activate() })
    $null = $form.ShowDialog()

    return $textBox.Text
}

# ========================================
# DNS Connection Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "DNS Connection Check" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

Wait-NetworkReady -Target $DNS -PingCount 2

Write-Host ""

# ========================================
# Domain Join Loop
# ========================================
$ErrorActionPreference = 'Stop'

while ($true) {

    # ========================================
    # Domain Join Process
    # ========================================
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Domain Join Process" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host ""

    Write-Host "Executing domain join: $DOMAIN / $USER" -ForegroundColor Yellow
    Write-Host ""

    try {
        # Create credentials
        $securePassword = ConvertTo-SecureString $PASS -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($USER, $securePassword)

        # Join domain
        Add-Computer -DomainName $DOMAIN -Credential $credential -Force

        Write-Host ""
        Show-Success "Domain join completed"
        Write-Host ""
        return (New-ModuleResult -Status "Success" -Message "Domain join completed")
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Host ""
        Show-Error "Domain join failed: $errorMsg"
        Write-Host ""

        $inputText = Show-ErrorDialog -Title "Domain Join Failed" -Message @"
Domain join failed:
$errorMsg

Possible causes:
  - System clock is out of sync with the domain
  - Computer name already exists in Active Directory
  - DNS name resolution failure (SRV / A record)
  - Domain controller is unreachable
  - Invalid credentials (username / password)
  - Insufficient permissions to join the domain
  - Network connectivity issue

Type 'adminstop' to abort and return to main menu.
"@

        if ($inputText -eq "adminstop") {
            Show-Info "Aborted by administrator (adminstop)"
            return (New-ModuleResult -Status "Error" -Message "Domain join failed (aborted by admin): $errorMsg")
        }

        Show-Info "Rechecking DNS connectivity before retry..."
        Write-Host ""
        Wait-NetworkReady -Target $DNS -PingCount 2
        Write-Host ""
        continue
    }
}
