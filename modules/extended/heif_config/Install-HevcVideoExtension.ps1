$HevcExtensions = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEVCVideoExtension"})

If (!$HevcExtensions) {
    try {
        Add-AppxProvisionedPackage -Online -PackagePath "$($PSScriptRoot)\Microsoft.HEVCVideoExtension_2.0.61301.0_neutral_~_8wekyb3d8bbwe.AppxBundle" -SkipLicense -ErrorAction Stop

        Show-Success "HEVC Video Extensions installed"
        return (New-ModuleResult -Status "Success" -Message "HEVC Video Extensions installed")
    }
    catch {
        Show-Error "Failed to install HEVC Video Extensions: $_"
        return (New-ModuleResult -Status "Error" -Message "Installation failed: $_")
    }
} Else {

    Show-Skip "HEVC Video Extensions already installed"
    return (New-ModuleResult -Status "Skipped" -Message "Already installed")
}