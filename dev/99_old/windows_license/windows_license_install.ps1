Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- 1. GUI Input Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Install Product Key"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(350, 20)
$label.Text = "Enter New Product Key (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX):"

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20, 50)
$textBox.Size = New-Object System.Drawing.Size(340, 20)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(180, 100)
$okButton.Size = New-Object System.Drawing.Size(80, 30)
$okButton.Text = "Install"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(270, 100)
$cancelButton.Size = New-Object System.Drawing.Size(80, 30)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton

$form.Controls.Add($label)
$form.Controls.Add($textBox)
$form.Controls.Add($okButton)
$form.Controls.Add($cancelButton)

# Show Dialog
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Operation Cancelled." -ForegroundColor Yellow
    exit
}

$newProductKey = $textBox.Text.Trim()

if ([string]::IsNullOrWhiteSpace($newProductKey)) {
    Write-Host "Error: Product Key cannot be empty." -ForegroundColor Red
    exit
}

# --- 2. Process Start ---
Write-Host "`n--- Starting Key Installation Process ---" -ForegroundColor Cyan

try {
    # Get the licensing service instance
    $service = Get-CimInstance -ClassName SoftwareLicensingService

    # --- 3. Uninstall Existing Key ---
    Write-Host "Checking for existing product key..." -ForegroundColor Gray
    
    # Get current Windows license that has a key
    $currentProduct = Get-CimInstance -ClassName SoftwareLicensingProduct | 
                      Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } | 
                      Select-Object -First 1

    if ($currentProduct) {
        Write-Host "Existing key found (Partial: $($currentProduct.PartialProductKey)). Uninstalling..." -ForegroundColor Yellow
        
        # UninstallProductKey requires the ID (GUID) of the product
        Invoke-CimMethod -InputObject $service -MethodName UninstallProductKey -Arguments @{ProductKeyID = $currentProduct.ID} | Out-Null
        
        Write-Host "Existing key uninstalled successfully." -ForegroundColor Green
    } else {
        Write-Host "No existing key found. Proceeding..." -ForegroundColor Gray
    }

    # --- 4. Install New Key ---
    Write-Host "Installing new product key: $newProductKey" -ForegroundColor Cyan
    Invoke-CimMethod -InputObject $service -MethodName InstallProductKey -Arguments @{ProductKey = $newProductKey} | Out-Null
    Write-Host "New key installed successfully." -ForegroundColor Green

    # NOTE: Activation (RefreshLicenseStatus) is intentionally skipped.

} catch {
    Write-Host "Error: An unexpected error occurred." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# --- 5. Verify Installation ---
Start-Sleep -Seconds 1
$finalCheck = Get-CimInstance -ClassName SoftwareLicensingProduct | 
              Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } | 
              Select-Object -First 1

Write-Host "`n--- Installation Result ---" -ForegroundColor White
if ($finalCheck) {
    Write-Host "Current Partial Key : $($finalCheck.PartialProductKey)"
    Write-Host "License Status      : $($finalCheck.LicenseStatus) (Not Activated)" -ForegroundColor Gray
} else {
    Write-Host "Warning: Key installation confirmed, but could not retrieve details immediately." -ForegroundColor Yellow
}
Write-Host "-------------------------"