$HevcAppxBundle = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEVCVideoExtensions"})

$HevcAppx = (Get-AppxPackage -AllUsers "Microsoft.HEVCVideoExtensions")

$HevcExt = (Get-AppxPackage "Microsoft.HEVCVideoExtensions")

$successCount = 0
$skipCount = 0
$failCount = 0

If ($HevcAppxBundle) {
    try {
        $HevcAppxBundle | Remove-AppxProvisionedPackage -Online -ErrorAction Stop

        Write-Host "The HEVC Video Extensions Store Package has been Successfully Unprovisioned on this System" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to unprovision HEVC Video Extensions: $_" -ForegroundColor Red
        $failCount++
    }
} Else {

    Write-Host "The HEVC Video Extensions Microsoft Store Package has Not been Provisioned on this System" -ForegroundColor Yellow
    $skipCount++
}

If ($HevcAppx) {
    try {
        $HevcAppx | Remove-AppxPackage -AllUsers -ErrorAction Stop

        Write-Host "The HEVC Video Extensions Microsoft Store Package has been Successfully Removed for All Users" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to remove HEVC Video Extensions for all users: $_" -ForegroundColor Red
        $failCount++
    }
} Else {

    Write-Host "The HEVC Video Extensions Microsoft Store Package is Not Installed for All Users" -ForegroundColor Yellow
    $skipCount++
}

If ($HevcExt) {
    try {
        $HevcExt | Remove-AppxPackage -ErrorAction Stop

        Write-Host "The HEVC Video Extensions Microsoft Store Package has been Successfully Removed for this User" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Failed to remove HEVC Video Extensions for current user: $_" -ForegroundColor Red
        $failCount++
    }
} Else {

    Write-Host "The HEVC Video Extensions Microsoft Store Package is Not Installed for this User" -ForegroundColor Yellow
    $skipCount++
}

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
