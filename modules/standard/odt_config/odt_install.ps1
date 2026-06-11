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
#                 and (Offline mode only) the Office\ offline source.
#                 Relative paths are resolved from the module root.
#                 If omitted, defaults to assets\.
#                 setup.exe is always loaded from assets\ regardless.
#   Mode        : (optional) "Offline" (default) or "Online".
#                  Offline: SourcePath is forcibly rewritten to AssetsFolder
#                           absolute path. Office\ offline source required.
#                  Online : SourcePath attribute is removed so ODT downloads
#                           from the Microsoft CDN. Office\ not required.
#                  If the column is absent or empty, Offline is assumed.
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

# Resolve install mode from entry (Offline default, optional Mode column)
function Get-EntryMode {
    param($Entry)
    if ($Entry.PSObject.Properties['Mode'] -and -not [string]::IsNullOrWhiteSpace($Entry.Mode)) {
        if ($Entry.Mode.Trim() -ieq "Online") { return "Online" }
    }
    return "Offline"
}

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

    $entryMode = Get-EntryMode -Entry $entry

    if ($xmlExists) {
        Write-Host "  $desc" -ForegroundColor Yellow
        Write-Host "    XML:    $($entry.XmlFileName)"
        Write-Host "    Assets: $entryAssetsDir"
        Write-Host "    Mode:   $entryMode"
    }
    else {
        Write-Host "  $desc [XML NOT FOUND]" -ForegroundColor Red
        Write-Host "    XML:    $($entry.XmlFileName)"
        Write-Host "    Assets: $entryAssetsDir"
        Write-Host "    Mode:   $entryMode"
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

# (f) Detect existing C2R Office. A clean machine proceeds; an existing
# install is judged by ProductReleaseIds (e.g. "O365BusinessRetail",
# comma-separated): if the detected set EXACTLY matches the union of
# Product IDs in this config's XMLs, the install already happened ->
# idempotent Skip. Anything else (partial overlap, extras, unparsable
# XML) keeps the original fail-closed abort - C2R products cannot
# coexist and an ambiguous state must not be skipped silently.
# Granularity is Product ID only (language/channel/edition differences
# are not detected here).
$c2rConfig = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
$productIds = if ($c2rConfig) { $c2rConfig.ProductReleaseIds } else { $null }
if (-not [string]::IsNullOrWhiteSpace($productIds)) {
    # Detected IDs: the Configuration value holds bare IDs (field-
    # confirmed on an M365 Apps for Business machine), but strip a
    # ".<digits>" version suffix defensively - the internal
    # ProductReleaseIDs key tree decorates names like
    # "O365BusinessRetail.16" and same-product false mismatches are
    # worse than the widened match.
    $detectedIds = @($productIds -split ',' | ForEach-Object { ($_.Trim() -replace '\.\d+$', '') } | Where-Object { $_ })

    # Target IDs: union of <Product ID> across every enabled entry's XML.
    # Parse failures contribute nothing -> empty target -> fail-closed.
    $targetIds = @()
    foreach ($e in $enabledEntries) {
        $eAssetsDir = if (-not [string]::IsNullOrWhiteSpace($e.AssetsFolder)) {
            if ([System.IO.Path]::IsPathRooted($e.AssetsFolder)) { $e.AssetsFolder } else { Join-Path $PSScriptRoot $e.AssetsFolder }
        } else { $AssetsDir }
        $eXmlPath = Join-Path $eAssetsDir $e.XmlFileName
        if (Test-Path $eXmlPath) {
            try {
                $eXml = [xml](Get-Content $eXmlPath -Encoding UTF8)
                foreach ($prod in @($eXml.Configuration.Add.Product)) {
                    if ($prod -and $prod.ID) { $targetIds += "$($prod.ID)".Trim() }
                }
            }
            catch { }
        }
    }
    $targetIds = @($targetIds | Where-Object { $_ } | Select-Object -Unique)

    $detectedKey = (@($detectedIds | Sort-Object) -join '|').ToLowerInvariant()
    $targetKey   = (@($targetIds   | Sort-Object) -join '|').ToLowerInvariant()

    if ($detectedKey -ne '' -and $detectedKey -eq $targetKey) {
        Show-Skip "Office already installed by this configuration: $productIds"
        Write-Host ""
        return (New-ModuleResult -Status "Skipped" -Message "Already installed: $productIds" -Verified $true)
    }

    Show-Error "Existing Click-to-Run Office detected: $productIds"
    Show-Error "Target products of this config: $(if ($targetIds.Count -gt 0) { $targetIds -join ',' } else { '(none parsed from XML)' })"
    Show-Error "Please uninstall existing Office before running ODT."
    Show-Info "Use SaRA tool (https://aka.ms/SaRA-officeUninstallFromPC) or manual uninstall."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Existing C2R Office detected: $productIds (target: $(if ($targetIds.Count -gt 0) { $targetIds -join ',' } else { 'none' }))")
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

    # Mark entry start time for ODT log collection window (see finally block).
    $entryStartTime = Get-Date

    try {
        # (a) Load XML and adjust <Add SourcePath> based on Mode
        $entryMode = Get-EntryMode -Entry $entry
        Show-Info "Preparing XML for $entryMode install..."
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

        if ($entryMode -ieq "Online") {
            # Online: remove SourcePath so ODT downloads from Microsoft CDN
            if ($AddNode.HasAttribute("SourcePath")) {
                $AddNode.RemoveAttribute("SourcePath")
            }
            Show-Info "Online mode: SourcePath removed (CDN download)"
        }
        else {
            # Offline: rewrite SourcePath to the absolute assets path
            $AddNode.SetAttribute("SourcePath", $entryAssetsDir)
            Show-Info "Offline mode: SourcePath set to $entryAssetsDir"
        }

        $XmlContent.Save($TempXmlPath)

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

        # Collect ODT logs to shared evidence bucket (.\evidence\odt_log\).
        # Filter: hostname-prefixed *.log files modified since this entry started.
        # ODT writes {COMPUTERNAME}-{yyyyMMdd}-{HHmm}[a-z].log; one entry can
        # produce multiple files (main log + correlation), so collect all fresh
        # matches instead of only the most recent one.
        $odtLogs = @(Get-ChildItem "C:\Windows\Temp" -Filter "$env:COMPUTERNAME-*.log" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -ge $entryStartTime })
        if ($odtLogs.Count -gt 0) {
            $logDest = Join-Path ".\evidence" "odt_log"
            if (-not (Test-Path $logDest)) {
                New-Item -Path $logDest -ItemType Directory -Force | Out-Null
            }

            # Prefix ensures uniqueness across sessions and PCs. Factory-default
            # or cloned images can share hostnames, which would collide on the
            # ODT-native filename alone. Session tag is derived from the evidence
            # base path's grandparent leaf (format: {ts}_{PCName}_{Serial}).
            $prefix = if (-not [string]::IsNullOrWhiteSpace($global:FabriqEvidenceBasePath)) {
                $evidenceRoot = Split-Path $global:FabriqEvidenceBasePath -Parent
                $leaf = Split-Path $evidenceRoot -Leaf
                ($leaf -replace '_evidence$', '')
            } else {
                "$(Get-Date -Format 'yyyy_MM_dd_HHmmss')_${env:COMPUTERNAME}"
            }

            foreach ($log in $odtLogs) {
                $destFile = Join-Path $logDest "${prefix}_$($log.Name)"
                Copy-Item $log.FullName $destFile -Force -ErrorAction SilentlyContinue
                Show-Info "ODT log collected: $destFile"
            }
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
