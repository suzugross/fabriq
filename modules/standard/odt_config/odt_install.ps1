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
