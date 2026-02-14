# 1. Trigger Activation (Refresh License Status)
Write-Host "Triggering Windows Activation via CIM..." -ForegroundColor Cyan

try {
    # FIX: Get the service instance first, then invoke the method on that object
    $service = Get-CimInstance -ClassName SoftwareLicensingService
    Invoke-CimMethod -InputObject $service -MethodName RefreshLicenseStatus | Out-Null
    
    Write-Host "Request sent. Waiting for status update..." -ForegroundColor Gray
}
catch {
    Write-Host "Error: Failed to trigger activation. details below:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    break
}

# Wait for the background process to settle
Start-Sleep -Seconds 3

# 2. Retrieve Windows License Information
$osLicense = Get-CimInstance -ClassName SoftwareLicensingProduct | 
             Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } | 
             Select-Object -First 1

# 3. Display Results
Write-Host "`n--- Activation Result ---" -ForegroundColor White

if ($null -eq $osLicense) {
    Write-Host "Error: No active Windows license found on this system." -ForegroundColor Red
}
else {
    switch ($osLicense.LicenseStatus) {
        0 { Write-Host "Status: [Unlicensed]" -ForegroundColor Red }
        1 { Write-Host "Status: [Licensed] - Success!" -ForegroundColor Green }
        2 { Write-Host "Status: [OOBE Grace Period]" -ForegroundColor Yellow }
        3 { Write-Host "Status: [Out of Tolerance]" -ForegroundColor Red }
        4 { Write-Host "Status: [Non-Genuine Grace Period]" -ForegroundColor Red }
        5 { Write-Host "Status: [Notification Mode]" -ForegroundColor Magenta }
        6 { Write-Host "Status: [Extended Grace Period]" -ForegroundColor Yellow }
        Default { Write-Host "Status: [Unknown] (Code: $($osLicense.LicenseStatus))" -ForegroundColor Gray }
    }
    
    $editionName = $osLicense.Name.Split(',')[0]
    Write-Host "Edition       : $editionName"
    Write-Host "Partial Key   : $($osLicense.PartialProductKey)"
}

Write-Host "-------------------------"