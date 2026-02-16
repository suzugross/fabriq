$HeifAppxBundle = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEIFImageExtension"})

$HeifAppx = (Get-AppxPackage -AllUsers "Microsoft.HEIFImageExtension")

$HeifExt = (Get-AppxPackage "Microsoft.HEIFImageExtension")

$successCount = 0
$skipCount = 0
$failCount = 0

If ($HeifAppxBundle) {
    try {
        $HeifAppxBundle | Remove-AppxProvisionedPackage -Online -ErrorAction Stop

        Show-Success "HEIF Image Extensions unprovisioned"
        $successCount++
    }
    catch {
        Show-Error "Failed to unprovision HEIF Image Extensions: $_"
        $failCount++
    }
} Else {

    Show-Skip "HEIF Image Extensions not provisioned"
    $skipCount++
}

If ($HeifAppx) {
    try {
        $HeifAppx | Remove-AppxPackage -AllUsers -ErrorAction Stop

        Show-Success "HEIF Image Extensions removed for all users"
        $successCount++
    }
    catch {
        Show-Error "Failed to remove HEIF Image Extensions for all users: $_"
        $failCount++
    }
} Else {

    Show-Skip "HEIF Image Extensions not installed for all users"
    $skipCount++
}

If ($HeifExt) {
    try {
        $HeifExt | Remove-AppxPackage -ErrorAction Stop

        Show-Success "HEIF Image Extensions removed for current user"
        $successCount++
    }
    catch {
        Show-Error "Failed to remove HEIF Image Extensions for current user: $_"
        $failCount++
    }
} Else {

    Show-Skip "HEIF Image Extensions not installed for current user"
    $skipCount++
}

# Return ModuleResult
$overallStatus = if ($failCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($successCount -gt 0 -and $failCount -gt 0) { "Partial" }
    elseif ($failCount -eq 0 -and $skipCount -gt 0) { "Skipped" }
    else { "Error" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $failCount")
