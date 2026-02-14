$HevcExtensions = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEVCVideoExtension"})

If (!$HevcExtensions) {
    try {
        Add-AppxProvisionedPackage -Online -PackagePath "$($PSScriptRoot)\Microsoft.HEVCVideoExtension_2.0.61301.0_neutral_~_8wekyb3d8bbwe.AppxBundle" -SkipLicense -ErrorAction Stop

        Write-Host "The HEVC Video Extensions Store App has been Installed to this System" -ForegroundColor Green
        return (New-ModuleResult -Status "Success" -Message "HEVC Video Extensions installed")
    }
    catch {
        Write-Host "[ERROR] Failed to install HEVC Video Extensions: $_" -ForegroundColor Red
        return (New-ModuleResult -Status "Error" -Message "Installation failed: $_")
    }
} Else {

    Write-Host "The HEVC Video Extensions Microsoft Store App is already Installed on this System" -ForegroundColor Yellow
    return (New-ModuleResult -Status "Skipped" -Message "Already installed")
}