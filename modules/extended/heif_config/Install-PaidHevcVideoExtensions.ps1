$PaidHevcExtensions = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEVCVideoExtension"})

If (!$PaidHevcExtensions) {
    try {
        Add-AppxProvisionedPackage -Online -PackagePath "$($PSScriptRoot)\Microsoft.HEVCVideoExtensions_2.0.60962.0_neutral_~_8wekyb3d8bbwe_(Paid).AppxBundle" -SkipLicense -ErrorAction Stop

        Write-Host "The Commercial HEVC Video Extensions Store App has been Installed to this System" -ForegroundColor Green
        return (New-ModuleResult -Status "Success" -Message "Commercial HEVC Video Extensions installed")
    }
    catch {
        Write-Host "[ERROR] Failed to install Commercial HEVC Video Extensions: $_" -ForegroundColor Red
        return (New-ModuleResult -Status "Error" -Message "Installation failed: $_")
    }
} Else {

    Write-Host "The Commercial HEVC Video Extensions Microsoft Store App is already Installed on this System" -ForegroundColor Yellow
    return (New-ModuleResult -Status "Skipped" -Message "Already installed")
}