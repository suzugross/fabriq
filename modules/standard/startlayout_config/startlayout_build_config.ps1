# ========================================
# Start Layout Build Script
# ========================================
# Reads an exported start layout JSON file, generates
# a customizations.xml for Windows Configuration Designer,
# and builds a provisioning package (.ppkg) using ICD.exe.
#
# [NOTES]
# - Requires Windows ADK (Configuration Designer component)
# - Input JSON must exist under json/ subdirectory
# - Output XML and PPKG are saved to xml/ and ppkg/ respectively
# ========================================

Write-Host ""
Show-Separator
Write-Host "Start Layout Build" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: CSV reading
# ========================================
$csvPath = Join-Path $PSScriptRoot "startlayout_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Id", "FileName")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load startlayout_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Pre-flight checks
# ========================================

# --- ICD.exe discovery ---
$icdExe = $null
$icdCandidates = @(
    "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Imaging and Configuration Designer\x86\ICD.exe"
    "$env:ProgramFiles\Windows Kits\10\Assessment and Deployment Kit\Imaging and Configuration Designer\x86\ICD.exe"
)

foreach ($candidate in $icdCandidates) {
    if (Test-Path $candidate) {
        $icdExe = $candidate
        break
    }
}

if (-not $icdExe) {
    $fromPath = Get-Command "ICD.exe" -ErrorAction SilentlyContinue
    if ($fromPath) {
        $icdExe = $fromPath.Source
    }
}

if (-not $icdExe) {
    Show-Error "ICD.exe not found. Install Windows ADK (Configuration Designer)."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "ICD.exe not found")
}

# --- StoreFile discovery (optional) ---
$storeFile = $null
$icdDir = Split-Path -Parent $icdExe
$storeFileCandidates = @(
    "Microsoft-Desktop-Provisioning.dat"
    "Microsoft-Common-Provisioning.dat"
)

foreach ($name in $storeFileCandidates) {
    $candidate = Join-Path $icdDir $name
    if (Test-Path $candidate) {
        $storeFile = $candidate
        break
    }
}

# --- Input JSON existence check ---
$jsonDir = Join-Path $PSScriptRoot "json"

foreach ($item in $enabledItems) {
    $jsonPath = Join-Path $jsonDir "$($item.FileName).json"
    if (-not (Test-Path $jsonPath)) {
        Show-Error "Input JSON not found: $jsonPath"
        Show-Error "Run Start Layout Backup first to generate the JSON file."
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Input JSON not found: $($item.FileName).json")
    }
}

# ========================================
# Step 3: Pre-execution display
# ========================================
$xmlDir  = Join-Path $PSScriptRoot "xml"
$ppkgDir = Join-Path $PSScriptRoot "ppkg"

Show-Info "ICD.exe: $icdExe"
if ($storeFile) {
    Show-Info "StoreFile: $storeFile"
}
else {
    Show-Warning "StoreFile not found. ICD.exe will use its default."
}
Write-Host ""

Show-Info "Build targets: $($enabledItems.Count) item(s)"
Write-Host ""

foreach ($item in $enabledItems) {
    $jsonPath = Join-Path $jsonDir "$($item.FileName).json"
    $xmlPath  = Join-Path $xmlDir  "$($item.FileName).xml"
    $ppkgPath = Join-Path $ppkgDir "$($item.FileName).ppkg"

    # Determine PPKG state
    if (Test-Path $ppkgPath) {
        $marker = "[REBUILD]"
        $markerColor = "Yellow"
    }
    else {
        $marker = "[NEW]"
        $markerColor = "White"
    }

    Write-Host "  [Id:$($item.Id)] $($item.FileName)  $marker" -ForegroundColor $markerColor
    Write-Host "    JSON: $jsonPath" -ForegroundColor DarkGray
    Write-Host "    XML:  $xmlPath" -ForegroundColor DarkGray
    Write-Host "    PPKG: $ppkgPath" -ForegroundColor DarkGray
    Write-Host ""
}

# ========================================
# Step 4: Confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Build provisioning package(s)?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Build execution
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

# Create output directories if they do not exist
foreach ($dir in @($xmlDir, $ppkgDir)) {
    if (-not (Test-Path $dir)) {
        try {
            $null = New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop
            Show-Info "Created directory: $dir"
        }
        catch {
            Show-Error "Failed to create directory: $dir - $_"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Failed to create directory: $dir")
        }
    }
}

# Build parameters (constants)
$packageVersion = "1.0"
$ownerType      = "OEM"
$rank           = 0

foreach ($item in $enabledItems) {
    $jsonPath = Join-Path $jsonDir "$($item.FileName).json"
    $xmlPath  = Join-Path $xmlDir  "$($item.FileName).xml"
    $ppkgPath = Join-Path $ppkgDir "$($item.FileName).ppkg"

    # ----------------------------------------
    # 5-1: Read and validate JSON
    # ----------------------------------------
    try {
        $jsonContent = Get-Content -Path $jsonPath -Raw -Encoding UTF8 -ErrorAction Stop
        $jsonObj = $jsonContent | ConvertFrom-Json
    }
    catch {
        Show-Error "Failed to read/parse JSON: $($item.FileName).json - $_"
        $failCount++
        Write-Host ""
        continue
    }

    if (-not $jsonObj.pinnedList) {
        Show-Warning "No 'pinnedList' key in JSON. Proceeding anyway."
    }
    else {
        $pinCount = @($jsonObj.pinnedList).Count
        Show-Info "pinnedList: $pinCount entries"
    }

    # ----------------------------------------
    # 5-2: Compact JSON and escape for XML
    # ----------------------------------------
    $jsonCompact = ($jsonObj | ConvertTo-Json -Depth 10 -Compress)
    $jsonForXml  = [System.Security.SecurityElement]::Escape($jsonCompact)
    $nameEscaped = [System.Security.SecurityElement]::Escape($item.FileName)

    # ----------------------------------------
    # 5-3: Generate customizations.xml
    # ----------------------------------------
    $packageId = [guid]::NewGuid().ToString()

    $xmlContent = @"
<?xml version="1.0" encoding="utf-8"?>
<WindowsCustomizations>
  <PackageConfig xmlns="urn:schemas-Microsoft-com:Windows-ICD-Package-Config.v1.0">
    <ID>{$packageId}</ID>
    <Name>$nameEscaped</Name>
    <Version>$packageVersion</Version>
    <OwnerType>$ownerType</OwnerType>
    <Rank>$rank</Rank>
    <Notes />
    <Description />
  </PackageConfig>
  <Settings xmlns="urn:schemas-microsoft-com:windows-provisioning">
    <Customizations>
      <Common>
        <Policies>
          <Start>
            <ConfigureStartPins>$jsonForXml</ConfigureStartPins>
          </Start>
        </Policies>
      </Common>
    </Customizations>
  </Settings>
</WindowsCustomizations>
"@

    try {
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($xmlPath, $xmlContent, $utf8Bom)
        Show-Success "Generated: $($item.FileName).xml"
    }
    catch {
        Show-Error "Failed to write XML: $($item.FileName).xml - $_"
        $failCount++
        Write-Host ""
        continue
    }

    # ----------------------------------------
    # 5-4: Build PPKG via ICD.exe
    # ----------------------------------------
    if (Test-Path $ppkgPath) {
        $null = Remove-Item $ppkgPath -Force -ErrorAction SilentlyContinue
    }

    $icdArgs = "/Build-ProvisioningPackage /CustomizationXML:`"$xmlPath`" /PackagePath:`"$ppkgPath`" +Overwrite"
    if ($storeFile) {
        $icdArgs += " /StoreFile:`"$storeFile`""
    }

    Show-Info "Executing ICD.exe..."
    Write-Host "  Command: `"$icdExe`" $icdArgs" -ForegroundColor DarkGray

    $icdOutput = cmd /c "`"$icdExe`" $icdArgs 2>&1"
    $icdExitCode = $LASTEXITCODE

    # Display ICD output
    if ($icdOutput) {
        foreach ($line in $icdOutput) {
            $lineStr = "$line"
            if ($lineStr -match "ERROR") {
                Write-Host "  $lineStr" -ForegroundColor Red
            }
            elseif ($lineStr -match "WARNING|WARN") {
                Write-Host "  $lineStr" -ForegroundColor Yellow
            }
            else {
                Write-Host "  $lineStr" -ForegroundColor DarkGray
            }
        }
    }

    # Strict validation: ExitCode AND file existence
    if ($icdExitCode -ne 0) {
        Show-Error "ICD.exe exited with code $icdExitCode"
        $failCount++
        Write-Host ""
        continue
    }

    if (-not (Test-Path $ppkgPath)) {
        Show-Error "PPKG file was not created: $ppkgPath"
        $failCount++
        Write-Host ""
        continue
    }

    $ppkgSize = (Get-Item $ppkgPath).Length
    if ($ppkgSize -eq 0) {
        Show-Error "PPKG file is empty: $ppkgPath"
        $failCount++
        Write-Host ""
        continue
    }

    Show-Success "Built: $($item.FileName).ppkg ($ppkgSize bytes)"
    Write-Host "  Path: $ppkgPath" -ForegroundColor DarkGray

    # Check for .cat file (generated alongside .ppkg)
    $catPath = $ppkgPath -replace '\.ppkg$', '.cat'
    if (Test-Path $catPath) {
        Write-Host "  Catalog: $catPath" -ForegroundColor DarkGray
    }

    $successCount++
    Write-Host ""
}

# ========================================
# Step 6: Result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Start Layout Build Results")
