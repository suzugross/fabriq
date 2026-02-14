$HeifAppxBundle = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEIFImageExtension"})

$HeifAppx = (Get-AppxPackage -AllUsers "Microsoft.HEIFImageExtension")

$HeifExt = (Get-AppxPackage "Microsoft.HEIFImageExtension")

$successCount = 0
$skipCount = 0
$failCount = 0

If ($HeifAppxBundle) {
    try {
        $HeifAppxBundle | Remove-AppxProvisionedPackage -Online -ErrorAction Stop

        Write-Host "The HEIF Image Extensions Microsoft Store Package has been Successfully Unprovisioned on this System" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to unprovision HEIF Image Extensions: $_" -ForegroundColor Red
        $failCount++
    }
} Else {

    Write-Host "The HEIF Image Extensions Microsoft Store Package has Not been Provisioned on this System" -ForegroundColor Yellow
    $skipCount++
}

If ($HeifAppx) {
    try {
        $HeifAppx | Remove-AppxPackage -AllUsers -ErrorAction Stop

        Write-Host "The HEIF Image Extensions Microsoft Store Package has been Successfully Removed for All Users" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to remove HEIF Image Extensions for all users: $_" -ForegroundColor Red
        $failCount++
    }
} Else {

    Write-Host "The HEIF Image Extensions Microsoft Store Package is Not Installed for All Users" -ForegroundColor Yellow
    $skipCount++
}

If ($HeifExt) {
    try {
        $HeifExt | Remove-AppxPackage -ErrorAction Stop

        Write-Host "The HEIF Image Extensions Microsoft Store Package has been Successfully Removed for this User" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to remove HEIF Image Extensions for current user: $_" -ForegroundColor Red
        $failCount++
    }
} Else {

    Write-Host "The HEIF Image Extensions Microsoft Store Package is Not Installed for this User" -ForegroundColor Yellow
    $skipCount++
}

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
