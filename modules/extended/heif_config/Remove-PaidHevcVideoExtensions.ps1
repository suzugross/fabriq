$HevcAppxBundle = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEVCVideoExtensions"})

$HevcAppx = (Get-AppxPackage -AllUsers "Microsoft.HEVCVideoExtensions")

$HevcExt = (Get-AppxPackage "Microsoft.HEVCVideoExtensions")

$successCount = 0
$skipCount = 0
$failCount = 0

If ($HevcAppxBundle) {
    try {
        $HevcAppxBundle | Remove-AppxProvisionedPackage -Online -ErrorAction Stop

        Show-Success "Commercial HEVC Video Extensions unprovisioned"
        $successCount++
    }
    catch {
        Show-Error "Failed to unprovision Commercial HEVC Video Extensions: $_"
        $failCount++
    }
} Else {

    Show-Skip "Commercial HEVC Video Extensions not provisioned"
    $skipCount++
}

If ($HevcAppx) {
    try {
        $HevcAppx | Remove-AppxPackage -AllUsers -ErrorAction Stop

        Show-Success "Commercial HEVC Video Extensions removed for all users"
        $successCount++
    }
    catch {
        Show-Error "Failed to remove Commercial HEVC Video Extensions for all users: $_"
        $failCount++
    }
} Else {

    Show-Skip "Commercial HEVC Video Extensions not installed for all users"
    $skipCount++
}

If ($HevcExt) {
    try {
        $HevcExt | Remove-AppxPackage -ErrorAction Stop

        Show-Success "Commercial HEVC Video Extensions removed for current user"
        $successCount++
    }
    catch {
        Show-Error "Failed to remove Commercial HEVC Video Extensions for current user: $_"
        $failCount++
    }
} Else {

    Show-Skip "Commercial HEVC Video Extensions not installed for current user"
    $skipCount++
}

return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount)
