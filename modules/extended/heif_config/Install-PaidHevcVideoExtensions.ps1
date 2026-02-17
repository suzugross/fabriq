$PaidHevcExtensions = (Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.HEVCVideoExtension"})

If (!$PaidHevcExtensions) {
    try {
        Add-AppxProvisionedPackage -Online -PackagePath "$($PSScriptRoot)\Microsoft.HEVCVideoExtensions_2.0.60962.0_neutral_~_8wekyb3d8bbwe_(Paid).AppxBundle" -SkipLicense -ErrorAction Stop

        Show-Success "Commercial HEVC Video Extensions installed"
        return (New-ModuleResult -Status "Success" -Message "Commercial HEVC Video Extensions installed")
    }
    catch {
        Show-Error "Failed to install Commercial HEVC Video Extensions: $_"
        return (New-ModuleResult -Status "Error" -Message "Installation failed: $_")
    }
} Else {

    Show-Skip "Commercial HEVC Video Extensions already installed"
    return (New-ModuleResult -Status "Skipped" -Message "Already installed")
}