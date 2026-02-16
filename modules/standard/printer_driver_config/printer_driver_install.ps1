# ========================================
# Printer Driver Installation Script
# ========================================

$INF_DIR = Join-Path $PSScriptRoot "INF"

Write-Host ""
Show-Separator
Write-Host "Printer Driver Installation" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# Check INF folder existence
if (-not (Test-Path $INF_DIR)) {
    Show-Error "INF folder not found: $INF_DIR"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "INF folder not found")
}

# Get subfolders in INF (assumed as model names)
$modelFolders = Get-ChildItem -Path $INF_DIR -Directory
if ($modelFolders.Count -eq 0) {
    Show-Error "No model folders found in INF directory"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No model folders found in INF directory")
}

# ========================================
# Step 1: Select Model
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "[Step 1] Select printer model to install" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

for ($i = 0; $i -lt $modelFolders.Count; $i++) {
    Write-Host "  [$($i + 1)] $($modelFolders[$i].Name)" -ForegroundColor White
}

Write-Host ""
Write-Host -NoNewline "Enter number (0 to cancel): "
$modelChoice = Read-Host

if ($modelChoice -eq '0') {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

$modelNum = 0
if (-not [int]::TryParse($modelChoice, [ref]$modelNum) -or $modelNum -lt 1 -or $modelNum -gt $modelFolders.Count) {
    Write-Host ""
    Show-Error "Invalid number"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid number")
}

$selectedFolder = $modelFolders[$modelNum - 1]
$modelName = $selectedFolder.Name

Write-Host ""
Show-Info "Selected: $modelName"
Write-Host ""

# ========================================
# Step 2: Search INF Files & Check Architecture
# ========================================
Show-Info "Searching for INF files..."

# Recursive search for INF in selected folder
$allInfFiles = Get-ChildItem -Path $selectedFolder.FullName -Recurse -Filter "*.inf"

if ($allInfFiles.Count -eq 0) {
    Show-Error "INF files not found"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "INF files not found")
}

# Determine current architecture
$arch = if ([Environment]::Is64BitOperatingSystem) { "NTamd64" } else { "NTx86" }

# Check architecture compatibility for each INF
$validInfFiles = @()

foreach ($inf in $allInfFiles) {
    $content = Get-Content -Path $inf.FullName -Encoding Default -ErrorAction SilentlyContinue
    if (-not $content) { continue }

    # Get model section names from Manufacturer section
    $inManufacturer = $false
    $modelSectionNames = @()

    foreach ($line in $content) {
        $trimmed = $line.Trim()

        # Detect [Manufacturer] section start
        if ($trimmed -match '^\[Manufacturer\]') {
            $inManufacturer = $true
            continue
        }

        # Break if moving to another section
        if ($inManufacturer -and $trimmed -match '^\[') {
            break
        }

        # Look for entries with architecture decoration in Manufacturer section
        if ($inManufacturer -and $trimmed -match $arch) {
            $modelSectionNames += $trimmed
        }
    }

    if ($modelSectionNames.Count -eq 0) { continue }

    # Check if models are defined in the corresponding architecture model section
    $hasModels = $false
    $modelNames = @()

    foreach ($line in $content) {
        $trimmed = $line.Trim()

        # Search for architecture-specific model section
        if ($trimmed -match "^\[.*\.$arch\]") {
            $inModelSection = $true
            continue
        }

        if ($inModelSection -and $trimmed -match '^\[') {
            $inModelSection = $false
            continue
        }

        # Extract "Model Name" = ... inside model section
        if ($inModelSection -and $trimmed -match '^"(.+?)"\s*=') {
            $foundModel = $Matches[1]
            if ($foundModel -notin $modelNames) {
                $modelNames += $foundModel
            }
            $hasModels = $true
        }
    }

    if ($hasModels) {
        $validInfFiles += [PSCustomObject]@{
            Path       = $inf.FullName
            Name       = $inf.Name
            RelPath    = $inf.FullName.Replace($selectedFolder.FullName + "\", "")
            ModelNames = $modelNames
        }
    }
}

if ($validInfFiles.Count -eq 0) {
    Show-Error "No INF files found for current architecture ($arch)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No INF files found for current architecture ($arch)")
}

# Select INF (Auto-select if only one)
$selectedInf = $null

if ($validInfFiles.Count -eq 1) {
    $selectedInf = $validInfFiles[0]
    Show-Info "INF File: $($selectedInf.Name) (Auto-selected)"
}
else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[Step 2] Select INF File" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    for ($i = 0; $i -lt $validInfFiles.Count; $i++) {
        $inf = $validInfFiles[$i]
        $models = $inf.ModelNames -join ", "
        Write-Host "  [$($i + 1)] $($inf.RelPath)" -ForegroundColor White
        Write-Host "         Models: $models" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host -NoNewline "Enter number (0 to cancel): "
    $infChoice = Read-Host

    if ($infChoice -eq '0') {
        Write-Host ""
        Show-Info "Canceled"
        Write-Host ""
        return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
    }

    $infNum = 0
    if (-not [int]::TryParse($infChoice, [ref]$infNum) -or $infNum -lt 1 -or $infNum -gt $validInfFiles.Count) {
        Write-Host ""
        Show-Error "Invalid number"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Invalid number")
    }

    $selectedInf = $validInfFiles[$infNum - 1]
}

Write-Host ""

# ========================================
# Confirm Installation
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "The following drivers will be installed" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Model Folder:  $modelName" -ForegroundColor White
Write-Host "  INF File:      $($selectedInf.Name)" -ForegroundColor White
Write-Host "  Target Models: $($selectedInf.ModelNames -join ', ')" -ForegroundColor White
Write-Host "  Architecture:  $arch" -ForegroundColor White
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not (Confirm-Execution -Message "Do you want to install?")) {
    Write-Host ""
    Show-Info "Canceled"
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Step 3: Register to Driver Store with pnputil
# ========================================
Show-Info "Registering to Driver Store..."

$pnpResult = & pnputil /add-driver "$($selectedInf.Path)" /install 2>&1
$pnpExitCode = $LASTEXITCODE
$pnpOutput = ($pnpResult | Out-String)

# Check if driver already exists in system
$alreadyExists = $pnpOutput -match 'already exists|既にシステムに存在'

if ($pnpExitCode -ne 0 -and -not $alreadyExists) {
    Show-Error "Failed to register driver with pnputil"
    foreach ($line in $pnpResult) {
        Write-Host "  $line" -ForegroundColor Gray
    }
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to register driver with pnputil")
}

if ($alreadyExists) {
    Show-Skip "Driver already exists in Driver Store"
}
else {
    Show-Success "Registered to Driver Store"
}
Write-Host ""

# ========================================
# Step 4: Resolve INF Path in Driver Store
# ========================================
Show-Info "Resolving Driver Store path..."

$infBaseName = $selectedInf.Name.ToLower() -replace '\.inf$', ''
$storeDir = Get-ChildItem "C:\WINDOWS\System32\DriverStore\FileRepository" -Directory -Filter "${infBaseName}.inf_amd64_*" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $storeDir) {
    Show-Error "INF not found in Driver Store"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "INF not found in Driver Store")
}

$storeInfPath = Join-Path $storeDir.FullName $selectedInf.Name

if (-not (Test-Path $storeInfPath)) {
    Show-Error "INF file not found in Driver Store: $storeInfPath"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "INF file not found in Driver Store")
}

Write-Host "[INFO] Store Path: $storeInfPath" -ForegroundColor Gray
Write-Host ""

# ========================================
# Step 5: Register Each Model with Add-PrinterDriver
# ========================================
$successCount = 0
$skipCount = 0
$errorCount = 0

foreach ($driverName in $selectedInf.ModelNames) {
    Show-Info "Registering printer driver: $driverName"

    # Check if driver already registered
    $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
    if ($existingDriver) {
        Show-Skip "Driver already registered: $driverName"
        $skipCount++
        continue
    }

    try {
        Add-PrinterDriver -Name $driverName -InfPath $storeInfPath -ErrorAction Stop
        Show-Success "Registration complete: $driverName"
        $successCount++
    }
    catch {
        Show-Error "Registration failed: $driverName - $_"
        $errorCount++
    }
}

Write-Host ""

# ========================================
# Result Summary
# ========================================
Show-Separator
Write-Host "Installation Results" -ForegroundColor Cyan
Show-Separator
if ($successCount -gt 0) {
    Write-Host "Success: $successCount items" -ForegroundColor Green
}
if ($skipCount -gt 0) {
    Write-Host "Skipped: $skipCount items (already installed)" -ForegroundColor Gray
}
if ($errorCount -gt 0) {
    Write-Host "Failed: $errorCount items" -ForegroundColor Red
}
Write-Host ""

# Return ModuleResult
$overallStatus = if ($errorCount -eq 0 -and $successCount -gt 0) { "Success" }
    elseif ($errorCount -eq 0 -and $skipCount -gt 0 -and $successCount -eq 0) { "Skipped" }
    elseif ($successCount -gt 0 -and $errorCount -gt 0) { "Partial" }
    elseif ($errorCount -gt 0) { "Error" }
    else { "Success" }
return (New-ModuleResult -Status $overallStatus -Message "Success: $successCount, Skip: $skipCount, Fail: $errorCount")