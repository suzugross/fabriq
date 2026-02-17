$HeifExtensions = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEIFImageExtension"})

If (!$HeifExtensions) {
    try {
        Add-AppxProvisionedPackage -Online -PackagePath "$($PSScriptRoot)\Microsoft.HEIFImageExtension_1.0.61171.0_neutral_~_8wekyb3d8bbwe.AppxBundle" -SkipLicense -ErrorAction Stop

        Show-Success "HEIF Image Extensions installed"
        return (New-ModuleResult -Status "Success" -Message "HEIF Image Extensions installed")
    }
    catch {
        Show-Error "Failed to install HEIF Image Extensions: $_"
        return (New-ModuleResult -Status "Error" -Message "Installation failed: $_")
    }
} Else {

    Show-Skip "HEIF Image Extensions already installed"
    return (New-ModuleResult -Status "Skipped" -Message "Already installed")
}