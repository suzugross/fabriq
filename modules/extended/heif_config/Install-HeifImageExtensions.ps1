$HeifExtensions = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEIFImageExtension"})

If (!$HeifExtensions) {
    try {
        Add-AppxProvisionedPackage -Online -PackagePath "$($PSScriptRoot)\Microsoft.HEIFImageExtension_1.0.61171.0_neutral_~_8wekyb3d8bbwe.AppxBundle" -SkipLicense -ErrorAction Stop

        Write-Host "The HEIF Image Extensions Store App has been Installed to this System" -ForegroundColor Green
        return (New-ModuleResult -Status "Success" -Message "HEIF Image Extensions installed")
    }
    catch {
        Write-Host "[ERROR] Failed to install HEIF Image Extensions: $_" -ForegroundColor Red
        return (New-ModuleResult -Status "Error" -Message "Installation failed: $_")
    }
} Else {

    Write-Host "The HEIF Image Extensions Microsoft Store App is already Installed on this System" -ForegroundColor Yellow
    return (New-ModuleResult -Status "Skipped" -Message "Already installed")
}