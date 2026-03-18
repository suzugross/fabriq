# ========================================
# Printer Driver Installation Script
# ========================================
# Supports two modes:
#   Auto mode:        Reads driver names from hostlist environment variables,
#                     scans INF files to find matching drivers, and installs automatically.
#   Interactive mode:  Falls back to manual folder/INF selection when no host is selected.
# ========================================

$INF_DIR = Join-Path $PSScriptRoot "INF"

# ========================================
# Helper Functions
# ========================================

function Get-ValidInfFiles {
    param(
        [string]$FolderPath,
        [string]$BasePath
    )
    # Determine current architecture
    $arch = if ([Environment]::Is64BitOperatingSystem) { "NTamd64" } else { "NTx86" }

    $allInfFiles = Get-ChildItem -Path $FolderPath -Recurse -Filter "*.inf"
    if ($allInfFiles.Count -eq 0) { return @() }

    $validInfFiles = @()

    foreach ($inf in $allInfFiles) {
        $content = Get-Content -Path $inf.FullName -Encoding Default -ErrorAction SilentlyContinue
        if (-not $content) { continue }

        # Get model section names from Manufacturer section
        $inManufacturer = $false
        $modelSectionNames = @()

        foreach ($line in $content) {
            $trimmed = $line.Trim()

            if ($trimmed -match '^\[Manufacturer\]') {
                $inManufacturer = $true
                continue
            }

            if ($inManufacturer -and $trimmed -match '^\[') {
                break
            }

            if ($inManufacturer -and $trimmed -match $arch) {
                $modelSectionNames += $trimmed
            }
        }

        if ($modelSectionNames.Count -eq 0) { continue }

        # Check if models are defined in the corresponding architecture model section
        $hasModels = $false
        $modelNames = @()
        $inModelSection = $false

        foreach ($line in $content) {
            $trimmed = $line.Trim()

            if ($trimmed -match "^\[.*\.$arch\]") {
                $inModelSection = $true
                continue
            }

            if ($inModelSection -and $trimmed -match '^\[') {
                $inModelSection = $false
                continue
            }

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
                RelPath    = $inf.FullName.Replace($BasePath + "\", "")
                ModelNames = $modelNames
            }
        }
    }

    return $validInfFiles
}

function Install-DriverFromInf {
    param(
        [PSCustomObject]$InfInfo,
        [string[]]$FilterDriverNames
    )

    $result = @{ Success = 0; Skip = 0; Fail = 0 }

    # --- Register to Driver Store with pnputil ---
    Show-Info "Registering to Driver Store: $($InfInfo.Name)"

    $pnpResult = & pnputil /add-driver "$($InfInfo.Path)" /install 2>&1
    $pnpExitCode = $LASTEXITCODE
    $pnpOutput = ($pnpResult | Out-String)

    $alreadyExists = $pnpOutput -match 'already exists|既にシステムに存在'

    if ($pnpExitCode -ne 0 -and -not $alreadyExists) {
        Show-Error "Failed to register driver with pnputil: $($InfInfo.Name)"
        foreach ($line in $pnpResult) {
            Write-Host "  $line" -ForegroundColor Gray
        }
        $result.Fail++
        return $result
    }

    if ($alreadyExists) {
        Show-Skip "Driver already exists in Driver Store"
    }
    else {
        Show-Success "Registered to Driver Store"
    }

    # --- Resolve INF Path in Driver Store ---
    $infBaseName = $InfInfo.Name.ToLower() -replace '\.inf$', ''
    $storeDir = Get-ChildItem "C:\WINDOWS\System32\DriverStore\FileRepository" -Directory -Filter "${infBaseName}.inf_amd64_*" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $storeDir) {
        Show-Error "INF not found in Driver Store"
        $result.Fail++
        return $result
    }

    $storeInfPath = Join-Path $storeDir.FullName $InfInfo.Name

    if (-not (Test-Path $storeInfPath)) {
        Show-Error "INF file not found in Driver Store: $storeInfPath"
        $result.Fail++
        return $result
    }

    Show-Info "Store Path: $storeInfPath"

    # --- Register Each Model with Add-PrinterDriver ---
    $targetModels = if ($FilterDriverNames -and $FilterDriverNames.Count -gt 0) {
        $filtered = $InfInfo.ModelNames | Where-Object { $_ -in $FilterDriverNames }
        if ($filtered.Count -gt 0) { $filtered } else { $InfInfo.ModelNames }
    }
    else {
        $InfInfo.ModelNames
    }

    foreach ($driverName in $targetModels) {
        Show-Info "Registering printer driver: $driverName"

        $existingDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
        if ($existingDriver) {
            Show-Skip "Driver already registered: $driverName"
            $result.Skip++
            continue
        }

        try {
            Add-PrinterDriver -Name $driverName -InfPath $storeInfPath -ErrorAction Stop
            Show-Success "Registration complete: $driverName"
            $result.Success++
        }
        catch {
            Show-Error "Registration failed: $driverName - $_"
            $result.Fail++
        }
    }

    return $result
}


# ========================================
# Main Script
# ========================================

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
# Mode Detection
# ========================================
$autoDriverNames = @()
for ($i = 1; $i -le 10; $i++) {
    $driverName = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_DRIVER")
    if (-not [string]::IsNullOrEmpty($driverName) -and $driverName -notin $autoDriverNames) {
        $autoDriverNames += $driverName
    }
}

$isAutoMode = ($autoDriverNames.Count -gt 0)


if ($isAutoMode) {
    # ========================================
    # AUTO MODE
    # ========================================
    Show-Info "Auto mode: $($autoDriverNames.Count) unique driver(s) from hostlist"
    Write-Host ""

    # Step 1: Scan all INF folders and build driver name mapping
    Show-Info "Scanning INF files for driver names..."

    $driverMap = @{}

    foreach ($folder in $modelFolders) {
        $validInfs = Get-ValidInfFiles -FolderPath $folder.FullName -BasePath $folder.FullName
        foreach ($inf in $validInfs) {
            foreach ($model in $inf.ModelNames) {
                if (-not $driverMap.ContainsKey($model)) {
                    $driverMap[$model] = [PSCustomObject]@{
                        InfInfo    = $inf
                        FolderName = $folder.Name
                    }
                }
            }
        }
    }

    Show-Info "Found $($driverMap.Count) driver(s) in INF files"
    Write-Host ""

    # Step 2: Match hostlist drivers against INF driver names
    $matchedDrivers = @()
    $unmatchedDrivers = @()

    foreach ($reqDriver in $autoDriverNames) {
        if ($driverMap.ContainsKey($reqDriver)) {
            $matchedDrivers += [PSCustomObject]@{
                DriverName = $reqDriver
                InfInfo    = $driverMap[$reqDriver].InfInfo
                FolderName = $driverMap[$reqDriver].FolderName
            }
        }
        else {
            $unmatchedDrivers += $reqDriver
        }
    }

    # Step 3: Display confirmation summary
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Auto Install: Printer Drivers" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    $idx = 1
    foreach ($m in $matchedDrivers) {
        Write-Host "  [$idx] $($m.DriverName)" -ForegroundColor White
        Write-Host "      Folder: $($m.FolderName)" -ForegroundColor DarkGray
        Write-Host "      INF:    $($m.InfInfo.Name)" -ForegroundColor DarkGray
        Write-Host ""
        $idx++
    }

    foreach ($u in $unmatchedDrivers) {
        Write-Host "  [!] $u -> No matching driver in INF files (SKIP)" -ForegroundColor Yellow
        Write-Host ""
    }

    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host ""

    if ($matchedDrivers.Count -eq 0) {
        Show-Warning "No matching drivers found in INF files"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "No matching drivers found in INF files")
    }

    # Step 4: Confirm execution
    $cancelResult = Confirm-ModuleExecution -Message "Do you want to install?"
    if ($null -ne $cancelResult) { return $cancelResult }

    Write-Host ""

    # Step 5: Install matched drivers
    $successCount = 0
    $skipCount = 0
    $failCount = 0

    # Deduplicate by INF path (multiple drivers may share the same INF)
    $processedInfs = @{}

    foreach ($m in $matchedDrivers) {
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "[Processing] $($m.DriverName)" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow

        $infPath = $m.InfInfo.Path

        if (-not $processedInfs.ContainsKey($infPath)) {
            $r = Install-DriverFromInf -InfInfo $m.InfInfo -FilterDriverNames @($m.DriverName)
            $processedInfs[$infPath] = $true
        }
        else {
            # INF already registered in driver store, only need Add-PrinterDriver
            $r = Install-DriverFromInf -InfInfo $m.InfInfo -FilterDriverNames @($m.DriverName)
        }

        $successCount += $r.Success
        $skipCount += $r.Skip
        $failCount += $r.Fail
        Write-Host ""
    }

    $failCount += $unmatchedDrivers.Count

    return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Title "Installation Results")
}
else {
    # ========================================
    # INTERACTIVE MODE (Original flow)
    # ========================================
    Show-Info "Interactive mode (no host selected)"
    Write-Host ""

    # Step 1: Select Model
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

    # Step 2: Search INF Files & Check Architecture
    Show-Info "Searching for INF files..."

    $validInfFiles = Get-ValidInfFiles -FolderPath $selectedFolder.FullName -BasePath $selectedFolder.FullName

    if ($validInfFiles.Count -eq 0) {
        $arch = if ([Environment]::Is64BitOperatingSystem) { "NTamd64" } else { "NTx86" }
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

    # Confirm Installation
    $arch = if ([Environment]::Is64BitOperatingSystem) { "NTamd64" } else { "NTx86" }

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

    $cancelResult = Confirm-ModuleExecution -Message "Do you want to install?"
    if ($null -ne $cancelResult) { return $cancelResult }

    Write-Host ""

    # Install
    $r = Install-DriverFromInf -InfInfo $selectedInf

    return (New-BatchResult -Success $r.Success -Skip $r.Skip -Fail $r.Fail -Title "Installation Results")
}
