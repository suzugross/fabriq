# ========================================
# Printer Driver Installation Script
# ========================================

$INF_DIR = Join-Path $PSScriptRoot "INF"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Printer Driver Installation" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check INF folder existence
if (-not (Test-Path $INF_DIR)) {
    Write-Host "[ERROR] INF folder not found: $INF_DIR" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "INF folder not found")
}

# Get subfolders in INF (assumed as model names)
$modelFolders = Get-ChildItem -Path $INF_DIR -Directory
if ($modelFolders.Count -eq 0) {
    Write-Host "[ERROR] No model folders found in INF directory" -ForegroundColor Red
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
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

$modelNum = 0
if (-not [int]::TryParse($modelChoice, [ref]$modelNum) -or $modelNum -lt 1 -or $modelNum -gt $modelFolders.Count) {
    Write-Host ""
    Write-Host "[ERROR] Invalid number" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid number")
}

$selectedFolder = $modelFolders[$modelNum - 1]
$modelName = $selectedFolder.Name

Write-Host ""
Write-Host "[INFO] Selected: $modelName" -ForegroundColor Cyan
Write-Host ""

# ========================================
# Step 2: Search INF Files & Check Architecture
# ========================================
Write-Host "[INFO] Searching for INF files..." -ForegroundColor Cyan

# Recursive search for INF in selected folder
$allInfFiles = Get-ChildItem -Path $selectedFolder.FullName -Recurse -Filter "*.inf"

if ($allInfFiles.Count -eq 0) {
    Write-Host "[ERROR] INF files not found" -ForegroundColor Red
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
    Write-Host "[ERROR] No INF files found for current architecture ($arch)" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No INF files found for current architecture ($arch)")
}

# Select INF (Auto-select if only one)
$selectedInf = $null

if ($validInfFiles.Count -eq 1) {
    $selectedInf = $validInfFiles[0]
    Write-Host "[INFO] INF File: $($selectedInf.Name) (Auto-selected)" -ForegroundColor Cyan
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
        Write-Host "[INFO] Canceled" -ForegroundColor Cyan
        Write-Host ""
        return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
    }

    $infNum = 0
    if (-not [int]::TryParse($infChoice, [ref]$infNum) -or $infNum -lt 1 -or $infNum -gt $validInfFiles.Count) {
        Write-Host ""
        Write-Host "[ERROR] Invalid number" -ForegroundColor Red
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
    Write-Host "[INFO] Canceled" -ForegroundColor Cyan
    Write-Host ""
    return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
}

Write-Host ""

# ========================================
# Step 3: Register to Driver Store with pnputil
# ========================================
Write-Host "[INFO] Registering to Driver Store..." -ForegroundColor Cyan

$pnpResult = & pnputil /add-driver "$($selectedInf.Path)" /install 2>&1
$pnpExitCode = $LASTEXITCODE
$pnpOutput = ($pnpResult | Out-String)

# Check if driver already exists in system
$alreadyExists = $pnpOutput -match 'already exists|既にシステムに存在'

if ($pnpExitCode -ne 0 -and -not $alreadyExists) {
    Write-Host "[ERROR] Failed to register driver with pnputil" -ForegroundColor Red
    foreach ($line in $pnpResult) {
        Write-Host "  $line" -ForegroundColor Gray
    }
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to register driver with pnputil")
}

if ($alreadyExists) {
    Write-Host "[SKIP] Driver already exists in Driver Store" -ForegroundColor Gray
}
else {
    Write-Host "[SUCCESS] Registered to Driver Store" -ForegroundColor Green
}
Write-Host ""

# ========================================
# Step 4: Resolve INF Path in Driver Store
# ========================================
Write-Host "[INFO] Resolving Driver Store path..." -ForegroundColor Cyan

$infBaseName = $selectedInf.Name.ToLower() -replace '\.inf$', ''
$storeDir = Get-ChildItem "C:\WINDOWS\System32\DriverStore\FileRepository" -Directory -Filter "${infBaseName}.inf_amd64_*" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $storeDir) {
    Write-Host "[ERROR] INF not found in Driver Store" -ForegroundColor Red
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "INF not found in Driver Store")
}

$storeInfPath = Join-Path $storeDir.FullName $selectedInf.Name

if (-not (Test-Path $storeInfPath)) {
    Write-Host "[ERROR] INF file not found in Driver Store: $storeInfPath" -ForegroundColor Red
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
    Write-Host "[INFO] Registering printer driver: $driverName" -ForegroundColor Cyan

    # Check if driver already registered
    $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
    if ($existingDriver) {
        Write-Host "[SKIP] Driver already registered: $driverName" -ForegroundColor Gray
        $skipCount++
        continue
    }

    try {
        Add-PrinterDriver -Name $driverName -InfPath $storeInfPath -ErrorAction Stop
        Write-Host "[SUCCESS] Registration complete: $driverName" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "[ERROR] Registration failed: $driverName - $_" -ForegroundColor Red
        $errorCount++
    }
}

Write-Host ""

# ========================================
# Result Summary
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Installation Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
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