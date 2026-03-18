# ========================================
# Office Deployment Tool (ODT) Installer
# ========================================
# Reads odt_list.csv to determine which XML
# configuration(s) to use, rewrites the <Add>
# SourcePath attribute to the resolved assets
# folder path, then invokes setup.exe /configure.
# Temp XML is always cleaned up via finally block.
#
# odt_list.csv columns:
#   Enabled     : 1 to enable
#   XmlFileName : ODT config XML filename (resolved under AssetsFolder)
#   Description : Display name
#   AssetsFolder: (optional) Per-entry assets folder containing XmlFileName
#                 and the Office\ offline source. Relative paths are resolved
#                 from the module root. If omitted, defaults to assets\.
#                 setup.exe is always loaded from assets\ regardless.
# ========================================

Write-Host ""
Show-Separator
Write-Host "ODT Install" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# 1. Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "odt_list.csv"

$allEntries = Import-ModuleCsv -Path $csvPath -RequiredColumns @("Enabled", "XmlFileName")
if ($null -eq $allEntries) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load odt_list.csv")
}

$enabledEntries = @($allEntries | Where-Object { $_.Enabled -eq "1" })

if ($enabledEntries.Count -eq 0) {
    Show-Info "No enabled entries in odt_list.csv"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# 2. Asset Path Setup
# ========================================
$AssetsDir   = Join-Path $PSScriptRoot "assets"
$SetupExePath = Join-Path $AssetsDir "setup.exe"

# ========================================
# 3. Pre-flight Check
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Installation List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$missingCount = 0

if (-not (Test-Path $SetupExePath)) {
    Show-Error "setup.exe not found: $SetupExePath"
    $missingCount++
}

foreach ($entry in $enabledEntries) {
    $desc           = if ($entry.Description) { $entry.Description } else { $entry.XmlFileName }
    $entryAssetsDir = if (-not [string]::IsNullOrWhiteSpace($entry.AssetsFolder)) {
        if ([System.IO.Path]::IsPathRooted($entry.AssetsFolder)) {
            $entry.AssetsFolder
        } else {
            Join-Path $PSScriptRoot $entry.AssetsFolder
        }
    } else {
        $AssetsDir
    }
    $xmlPath   = Join-Path $entryAssetsDir $entry.XmlFileName
    $xmlExists = Test-Path $xmlPath

    if ($xmlExists) {
        Write-Host "  $desc" -ForegroundColor Yellow
        Write-Host "    XML:    $($entry.XmlFileName)"
        Write-Host "    Assets: $entryAssetsDir"
    }
    else {
        Write-Host "  $desc [XML NOT FOUND]" -ForegroundColor Red
        Write-Host "    XML:    $($entry.XmlFileName)"
        Write-Host "    Assets: $entryAssetsDir"
        $missingCount++
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

if (-not (Test-Path $SetupExePath)) {
    Show-Error "setup.exe is missing. Cannot proceed."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "setup.exe not found in assets")
}

if ($missingCount -gt 0) {
    Show-Warning "$missingCount item(s) with missing XML will be skipped"
    Write-Host ""
}

# ========================================
# 3.5 Environment Pre-check & Cleanup
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Environment Pre-check" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# (a) Stop running Office processes to prevent C2R lock conflicts
$officeProcesses = @(
    "WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSPUB",
    "MSACCESS", "VISIO", "LYNC", "Teams", "OfficeClickToRun", "OfficeC2RClient"
)
foreach ($procName in $officeProcesses) {
    $running = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($running) {
        Show-Warning "Running Office process detected: $procName (PID: $($running.Id -join ', '))"
        Stop-Process -Name $procName -Force -ErrorAction SilentlyContinue
        Show-Info "Office process stopped: $procName"
    }
}

# (b) Stop ClickToRunSvc service to release C2R locks
$c2rService = Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue
if ($c2rService -and $c2rService.Status -eq "Running") {
    Stop-Service -Name "ClickToRunSvc" -Force -ErrorAction SilentlyContinue
    Show-Info "ClickToRunSvc stopped"
}

# (c) Remove Store-based Office AppX packages (follows storeapp_config pattern)
$storeOfficePatterns = @("*OneNote*", "*Office.Desktop*", "*Office.OneNote*", "*OfficeSway*")
foreach ($pattern in $storeOfficePatterns) {
    # Current user packages
    $appxPkgs = Get-AppxPackage $pattern -ErrorAction SilentlyContinue
    foreach ($pkg in $appxPkgs) {
        Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction SilentlyContinue
        Show-Info "Removed Store app (User): $($pkg.Name)"
    }
    # Provisioned packages (prevents auto-install for new users)
    $provPkgs = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like $pattern }
    foreach ($provPkg in $provPkgs) {
        Remove-AppxProvisionedPackage -Online -PackageName $provPkg.PackageName -ErrorAction SilentlyContinue
        Show-Info "Removed Store app (Provisioned): $($provPkg.DisplayName)"
    }
}

# (d) Ensure Windows Installer service is not disabled
$msiService = Get-Service -Name "msiserver" -ErrorAction SilentlyContinue
if ($msiService -and $msiService.StartType -eq "Disabled") {
    Set-Service -Name "msiserver" -StartupType Manual
    Show-Warning "Windows Installer was disabled. Changed to Manual."
}

# (e) Check disk space on system drive
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive)'" -ErrorAction SilentlyContinue
if ($disk) {
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    if ($freeGB -lt 10) {
        Show-Warning "Low disk space: ${freeGB}GB (10GB+ recommended)"
    } else {
        Show-Info "Disk space: ${freeGB}GB"
    }
}

# (f) Detect existing C2R Office — abort if found (cannot coexist)
$c2rConfig = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
$productIds = if ($c2rConfig) { $c2rConfig.ProductReleaseIds } else { $null }
if (-not [string]::IsNullOrWhiteSpace($productIds)) {
    Show-Error "Existing Click-to-Run Office detected: $productIds"
    Show-Error "Please uninstall existing Office before running ODT."
    Show-Info "Use SaRA tool (https://aka.ms/SaRA-officeUninstallFromPC) or manual uninstall."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Existing C2R Office detected: $productIds")
}
Show-Info "No existing C2R Office detected. Environment is clean."

Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

# ========================================
# 4. Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with Office installation?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# 5. Execute per Entry
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($entry in $enabledEntries) {
    $desc           = if ($entry.Description) { $entry.Description } else { $entry.XmlFileName }
    $entryAssetsDir = if (-not [string]::IsNullOrWhiteSpace($entry.AssetsFolder)) {
        if ([System.IO.Path]::IsPathRooted($entry.AssetsFolder)) {
            $entry.AssetsFolder
        } else {
            Join-Path $PSScriptRoot $entry.AssetsFolder
        }
    } else {
        $AssetsDir
    }
    $ConfigXmlPath = Join-Path $entryAssetsDir $entry.XmlFileName
    $TempXmlPath   = Join-Path $env:TEMP "fabriq_odt_$(Get-Date -Format 'yyyyMMddHHmmss').xml"

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Installing: $desc" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Skip if XML missing
    if (-not (Test-Path $ConfigXmlPath)) {
        Show-Skip "XML not found: $($entry.XmlFileName)"
        Write-Host ""
        $skipCount++
        continue
    }

    try {
        # (a) Load XML and rewrite <Add SourcePath> to absolute assets path
        Show-Info "Rewriting SourcePath in $($entry.XmlFileName)..."
        $XmlContent = [xml](Get-Content $ConfigXmlPath -Encoding UTF8)

        if ($null -eq $XmlContent.Configuration) {
            Show-Error "No <Configuration> node found in $($entry.XmlFileName)"
            $failCount++
            continue
        }

        $AddNode = $XmlContent.Configuration.Add
        if ($null -eq $AddNode) {
            Show-Error "No <Add> node found in $($entry.XmlFileName)"
            $failCount++
            continue
        }

        $AddNode.SetAttribute("SourcePath", $entryAssetsDir)
        $XmlContent.Save($TempXmlPath)
        Show-Info "SourcePath set to: $entryAssetsDir"

        # (b) Execute setup.exe /configure
        Show-Info "Starting setup.exe. This may take several minutes..."
        Write-Host ""

        $Arguments = "/configure `"$TempXmlPath`""
        $proc = Start-Process -FilePath $SetupExePath `
            -ArgumentList $Arguments `
            -Wait -NoNewWindow -PassThru

        # (c) Evaluate exit code
        if ($proc.ExitCode -eq 0) {
            Show-Success "$desc installed successfully (ExitCode: 0)"
            $successCount++
        }
        else {
            Show-Error "$desc completed with ExitCode: $($proc.ExitCode)"
            Show-Info "Check ODT logs in C:\Windows\Temp for details"
            $failCount++
        }
    }
    catch {
        Show-Error "$desc : $($_.Exception.Message)"
        $failCount++
    }
    finally {
        if (Test-Path $TempXmlPath) {
            Remove-Item -Path $TempXmlPath -Force -ErrorAction SilentlyContinue
        }

        # Collect ODT log to evidence path
        $odtLog = Get-ChildItem "C:\Windows\Temp" -Filter "SetupExe(*.log)" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($odtLog) {
            $logDest = if ($global:FabriqEvidenceBasePath) {
                Join-Path $global:FabriqEvidenceBasePath "odt_logs"
            } else {
                Join-Path ".\evidence" "odt_logs"
            }
            if (-not (Test-Path $logDest)) {
                New-Item -Path $logDest -ItemType Directory -Force | Out-Null
            }
            $destFile = Join-Path $logDest $odtLog.Name
            Copy-Item $odtLog.FullName $destFile -Force -ErrorAction SilentlyContinue
            Show-Info "ODT log collected: $destFile"
        }
    }

    Write-Host ""
}

# ========================================
# 6. Result
# ========================================
$total = $enabledEntries.Count

if ($failCount -gt 0 -and $successCount -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "ODT Install failed ($failCount/$total)")
}
elseif ($failCount -gt 0) {
    return (New-ModuleResult -Status "Partial" -Message "ODT Install partial ($successCount ok, $failCount failed, $skipCount skipped)")
}
elseif ($skipCount -eq $total) {
    return (New-ModuleResult -Status "Skipped" -Message "All entries skipped (missing XML)")
}
else {
    return (New-ModuleResult -Status "Success" -Message "ODT Install complete ($successCount/$total)")
}
