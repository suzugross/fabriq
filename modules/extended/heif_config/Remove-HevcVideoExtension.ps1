$HevcAppxBundle = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEVCVideoExtension"})

$HevcAppx = (Get-AppxPackage -AllUsers "Microsoft.HEVCVideoExtension")

$HevcExt = (Get-AppxPackage "Microsoft.HEVCVideoExtension")

$successCount = 0
$skipCount = 0
$failCount = 0

If ($HevcAppxBundle) {
    try {
        $HevcAppxBundle | Remove-AppxProvisionedPackage -Online -ErrorAction Stop

        Show-Success "HEVC Video Extensions unprovisioned"
        $successCount++
    }
    catch {
        Show-Error "Failed to unprovision HEVC Video Extensions: $_"
        $failCount++
    }
} Else {

    Show-Skip "HEVC Video Extensions not provisioned"
    $skipCount++
}

If ($HevcAppx) {
    try {
        $HevcAppx | Remove-AppxPackage -AllUsers -ErrorAction Stop

        Show-Success "HEVC Video Extensions removed for all users"
        $successCount++
    }
    catch {
        Show-Error "Failed to remove HEVC Video Extensions for all users: $_"
        $failCount++
    }
} Else {

    Show-Skip "HEVC Video Extensions not installed for all users"
    $skipCount++
}

If ($HevcExt) {
    try {
        $HevcExt | Remove-AppxPackage -ErrorAction Stop

        Show-Success "HEVC Video Extensions removed for current user"
        $successCount++
    }
    catch {
        Show-Error "Failed to remove HEVC Video Extensions for current user: $_"
        $failCount++
    }
} Else {

    Show-Skip "HEVC Video Extensions not installed for current user"
    $skipCount++
}

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
