# ========================================
# App Inventory Export Script
# ========================================
# Scans the HKLM Uninstall registry hives for installed
# legacy desktop applications and exports the results to
# a CSV file under evidence\inventory\.
# The exported CSV is designed to be used directly as
# input for a bloatware removal module (Phase 2).
#
# NOTES:
# - Administrator privileges are required.
# - UWP / AppxPackage apps are excluded (no registry
#   uninstall entry exists for them).
# - HKCU is excluded; OEM bloatware is system-wide.
# ========================================

Write-Host ""
Show-Separator
Write-Host "App Inventory Export" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 2: Pre-flight check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required to read the Uninstall registry."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}


# ========================================
# Step 3: Pre-execution display
# ========================================
$pcName    = if (-not [string]::IsNullOrEmpty($env:SELECTED_NEW_PCNAME)) { $env:SELECTED_NEW_PCNAME } else { $env:COMPUTERNAME }
$dateStr   = Get-Date -Format "yyyy_MM_dd_HHmmss"
$outputDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\..\..\evidence\inventory"))
$outPath   = Join-Path $outputDir "app_inventory_${dateStr}_${pcName}.csv"

$regPaths = @(
    [PSCustomObject]@{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*";             Arch = "64bit" }
    [PSCustomObject]@{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"; Arch = "32bit" }
)

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Scan Targets" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

foreach ($rp in $regPaths) {
    Write-Host "  [$($rp.Arch)]" -ForegroundColor Yellow
    Write-Host "    $($rp.Path)" -ForegroundColor White
    Write-Host ""
}

Write-Host "  Output:" -ForegroundColor Cyan
Write-Host "    $outPath" -ForegroundColor White
Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""


# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Export installed app list to CSV?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Scan and export
# ========================================
$allApps    = @()
$seenNames  = @{}
$scanFailed = $false

foreach ($rp in $regPaths) {
    Show-Info "Scanning $($rp.Arch) registry..."
    try {
        $entries = @(Get-ItemProperty $rp.Path -ErrorAction SilentlyContinue)
        foreach ($entry in $entries) {
            if ([string]::IsNullOrWhiteSpace($entry.DisplayName)) { continue }
            if ($seenNames.ContainsKey($entry.DisplayName))       { continue }
            $seenNames[$entry.DisplayName] = $true

            $allApps += [PSCustomObject]@{
                Enabled              = "0"
                DisplayName          = $entry.DisplayName
                Publisher            = if ($entry.Publisher)            { $entry.Publisher }            else { "" }
                DisplayVersion       = if ($entry.DisplayVersion)       { $entry.DisplayVersion }       else { "" }
                Architecture         = $rp.Arch
                WindowsInstaller     = if ($entry.WindowsInstaller)     { "$($entry.WindowsInstaller)" } else { "0" }
                QuietUninstallString = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { "" }
                UninstallString      = if ($entry.UninstallString)      { $entry.UninstallString }      else { "" }
                NoRemove             = if ($entry.NoRemove)             { "$($entry.NoRemove)" }        else { "0" }
                SystemComponent      = if ($entry.SystemComponent)      { "$($entry.SystemComponent)" } else { "0" }
                InstallDate          = if ($entry.InstallDate)          { $entry.InstallDate }          else { "" }
                RegistryKey          = $entry.PSPath -replace "Microsoft.PowerShell.Core\\Registry::", ""
            }
        }
    }
    catch {
        Show-Warning "Failed to scan $($rp.Arch) hive: $_"
        $scanFailed = $true
    }
}

$allApps = @($allApps | Sort-Object Publisher, DisplayName)

$sysCompCount  = @($allApps | Where-Object { $_.SystemComponent -eq "1" }).Count
$noRemoveCount = @($allApps | Where-Object { $_.NoRemove -eq "1" }).Count

Write-Host ""

# Create output directory if needed
try {
    if (-not (Test-Path $outputDir)) {
        $null = New-Item -ItemType Directory -Path $outputDir -Force
    }
}
catch {
    Show-Error "Failed to create output directory: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to create output directory: $_")
}

# Export CSV
try {
    $allApps | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8 -Force
    Show-Success "Exported $($allApps.Count) apps to CSV"
    Show-Info "Output: $outPath"
    Write-Host ""
    if ($sysCompCount -gt 0) {
        Show-Warning "$sysCompCount app(s) flagged SystemComponent=1 (do not remove without verification)"
    }
    if ($noRemoveCount -gt 0) {
        Show-Warning "$noRemoveCount app(s) flagged NoRemove=1 (cannot be removed conventionally)"
    }
}
catch {
    Show-Error "Failed to export CSV: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to export CSV: $_")
}

Write-Host ""


# ========================================
# Step 6: Result
# ========================================
$msg = "Exported $($allApps.Count) apps to app_inventory_${dateStr}_${pcName}.csv"
if ($scanFailed) {
    return (New-ModuleResult -Status "Partial" -Message "$msg (some hives could not be read)")
}
return (New-ModuleResult -Status "Success" -Message $msg)
