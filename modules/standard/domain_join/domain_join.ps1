# ========================================
# Domain Join Script
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host "Executing domain join process..." -ForegroundColor Cyan
Write-Host ""

# ========================================
# Load domain.csv
# ========================================
$csvPath = Join-Path $PSScriptRoot "domain.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "[ERROR] domain.csv not found: $csvPath" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "domain.csv not found")
}

try {
    $domainList = @(Import-Csv -Path $csvPath -Encoding Default)
}
catch {
    Write-Host "[ERROR] Failed to load domain.csv: $_" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "Failed to load domain.csv: $_")
}

if ($domainList.Count -eq 0) {
    Write-Host "[ERROR] domain.csv contains no data" -ForegroundColor Red
    return (New-ModuleResult -Status "Error" -Message "domain.csv contains no data")
}

# Note: CSV headers must match these keys exactly
$domainEntry = $domainList[0]
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
# DNS Ping + Domain Join Loop
# ========================================
while ($true) {

    # ========================================
    # Check DNS Connection
    # ========================================
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "DNS Connection Check" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host ""

    Write-Host "[INFO] Checking connection to DNS server ($DNS)..." -ForegroundColor Cyan

    $dnsOk = $false
    try {
        $pingResult = Test-Connection -ComputerName $DNS -Count 2 -Quiet
        if ($pingResult) {
            $dnsOk = $true
        }
    }
    catch {
        $dnsOk = $false
    }

    if (-not $dnsOk) {
        Write-Host "[ERROR] Ping to DNS server failed" -ForegroundColor Red

        $inputText = Show-ErrorDialog -Title "DNS Connection Failed" -Message @"
Ping to DNS server ($DNS) failed.

Please check the following and press Retry:
  - LAN cable is securely connected
  - Network adapter is enabled
  - Correct VLAN / network segment
  - DNS server ($DNS) is reachable from this network

Type 'adminstop' to abort and return to main menu.
"@

        if ($inputText -eq "adminstop") {
            Write-Host "[INFO] Aborted by administrator (adminstop)" -ForegroundColor Yellow
            return (New-ModuleResult -Status "Error" -Message "Aborted by administrator (adminstop)")
        }

        Write-Host "[INFO] Retrying DNS connection..." -ForegroundColor Cyan
        Write-Host ""
        continue
    }

    Write-Host "[SUCCESS] Ping to DNS server succeeded" -ForegroundColor Green
    Write-Host ""

    # ========================================
    # Domain Join Process
    # ========================================
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Domain Join Process" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host ""

    Write-Host "Executing domain join: $DOMAIN / $USER" -ForegroundColor Yellow
    Write-Host ""

    $ErrorActionPreference = 'Stop'

    try {
        # Create credentials
        $securePassword = ConvertTo-SecureString $PASS -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($USER, $securePassword)

        # Join domain
        Add-Computer -DomainName $DOMAIN -Credential $credential -Force

        Write-Host ""
        Write-Host "[SUCCESS] Domain join completed" -ForegroundColor Green
        Write-Host ""
        return (New-ModuleResult -Status "Success" -Message "Domain join completed")
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Host ""
        Write-Host "[ERROR] Domain join failed: $errorMsg" -ForegroundColor Red
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
            Write-Host "[INFO] Aborted by administrator (adminstop)" -ForegroundColor Yellow
            return (New-ModuleResult -Status "Error" -Message "Domain join failed (aborted by admin): $errorMsg")
        }

        Write-Host "[INFO] Returning to DNS connection check..." -ForegroundColor Cyan
        Write-Host ""
        continue
    }
}
