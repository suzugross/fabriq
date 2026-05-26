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

    # exit 259 (ERROR_NO_MORE_ITEMS) is pnputil's locale-independent signal for
    # "package already in driver store, nothing to add" — treat as success.
    $pnpAlreadyInStore = $alreadyExists -or ($pnpExitCode -eq 259)
    $pnpOk = ($pnpExitCode -eq 0) -or $pnpAlreadyInStore

    if (-not $pnpOk) {
        Show-Error "Failed to register driver with pnputil: $($InfInfo.Name)"
        foreach ($line in $pnpResult) {
            Show-Warning "  $line"
        }
        $result.Fail++
        return $result
    }

    if ($pnpAlreadyInStore) {
        Show-Skip "Driver already exists in Driver Store"
    }
    else {
        Show-Success "Registered to Driver Store"
    }

    # --- Resolve INF Path in Driver Store (best effort) ---
    # If the resolution fails (e.g. a pre-staged package occupies a folder named
    # after a different INF basename), we do NOT early-return. Strategy 2 in the
    # model loop below will rescue via -Name only.
    $infBaseName = $InfInfo.Name.ToLower() -replace '\.inf$', ''
    $storeDir = Get-ChildItem "C:\WINDOWS\System32\DriverStore\FileRepository" -Directory -Filter "${infBaseName}.inf_amd64_*" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    $storeInfPath = $null
    if ($storeDir) {
        $candidatePath = Join-Path $storeDir.FullName $InfInfo.Name
        if (Test-Path $candidatePath) {
            $storeInfPath = $candidatePath
            Show-Info "Store Path: $storeInfPath"
        }
        else {
            Show-Info "INF file not present at expected store path: $candidatePath"
            Show-Info "Will use -Name only fallback for each model"
        }
    }
    else {
        Show-Info "Store folder for '$infBaseName' not found in DriverStore FileRepository"
        Show-Info "Will use -Name only fallback (driver may be in DriverStore under a different INF basename)"
    }

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

        $registered = $false

        # Strategy 1: -Name -InfPath (preferred — pins registration to a specific INF in the store)
        if ($storeInfPath) {
            try {
                Add-PrinterDriver -Name $driverName -InfPath $storeInfPath -ErrorAction Stop
                Show-Success "Registration complete: $driverName"
                $result.Success++
                $registered = $true
            }
            catch {
                Show-Warning "Add-PrinterDriver -InfPath failed for ${driverName}: $($_.Exception.Message)"
                Show-Info "Falling back to -Name only"
            }
        }

        # Strategy 2: -Name only (rescues from pre-staged DriverStore packages whose
        # repository folder name differs from this INF's basename — e.g. drivers
        # placed by Windows Update or OEM image)
        if (-not $registered) {
            try {
                Add-PrinterDriver -Name $driverName -ErrorAction Stop
                Show-Success "Registration complete (from DriverStore global): $driverName"
                $result.Success++
                $registered = $true
            }
            catch {
                Show-Error "Registration failed: $driverName - $($_.Exception.Message)"
                $result.Fail++
            }
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

# ========================================
# Step 0: Auto-extract EXE/ZIP archives in INF/
# ========================================
# If INF/ contains .exe or .zip files, extract them to same-named folders
# using the bundled 7z.exe. Already-extracted folders are skipped (idempotent).
# If 7z.exe is not available, .zip files fall back to Expand-Archive.
$7zPath = Join-Path $PSScriptRoot "tools\7z.exe"
$has7z = Test-Path $7zPath

$archives = @(Get-ChildItem -Path $INF_DIR -File | Where-Object {
    $_.Extension -in @('.exe', '.zip')
})

if ($archives.Count -gt 0) {
    Show-Info "Found $($archives.Count) archive(s) in INF/"

    if (-not $has7z) {
        Show-Warning "7z.exe not found in tools/ - cannot extract .exe archives"
        Show-Info "Place 7z.exe + 7z.dll in: $(Join-Path $PSScriptRoot 'tools')"

        foreach ($arc in $archives) {
            if ($arc.Extension -eq '.zip') {
                $targetDir = Join-Path $INF_DIR ($arc.BaseName)
                if (Test-Path $targetDir) {
                    Show-Skip "Already extracted: $($arc.Name)"
                    continue
                }
                Show-Info "Extracting (Expand-Archive): $($arc.Name)"
                try {
                    Expand-Archive -Path $arc.FullName -DestinationPath $targetDir -Force -ErrorAction Stop
                    Show-Success "Extracted: $($arc.Name) -> $($arc.BaseName)/"
                }
                catch {
                    Show-Warning "Failed to extract $($arc.Name): $_"
                }
            }
        }
    }
    else {
        foreach ($arc in $archives) {
            $targetDir = Join-Path $INF_DIR ($arc.BaseName)
            if (Test-Path $targetDir) {
                Show-Skip "Already extracted: $($arc.Name)"
                continue
            }
            Show-Info "Extracting: $($arc.Name)"
            $null = & $7zPath x $arc.FullName "-o$targetDir" -y -bso0 -bsp0 2>&1
            if ($LASTEXITCODE -eq 0) {
                Show-Success "Extracted: $($arc.Name) -> $($arc.BaseName)/"
            }
            else {
                Show-Warning "Failed to extract $($arc.Name) (exit code: $LASTEXITCODE)"
                if (Test-Path $targetDir) {
                    Remove-Item -Path $targetDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    Write-Host ""
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
# Collect driver names from two sources (unioned):
#   1. Hostlist environment variables (SELECTED_PRINTER_1..10_DRIVER)
#   2. printer_driver_list.csv (optional, module-local, TargetHost filtered)
$autoDriverNames = @()

for ($i = 1; $i -le 10; $i++) {
    $driverName = [Environment]::GetEnvironmentVariable("SELECTED_PRINTER_${i}_DRIVER")
    if (-not [string]::IsNullOrEmpty($driverName) -and $driverName -notin $autoDriverNames) {
        $autoDriverNames += $driverName
    }
}

# Collect from printer_driver_list.csv (optional)
# TargetHost column: empty = all hosts, value = exact match with SELECTED_NEW_PCNAME.
$driverCsvPath = Join-Path $PSScriptRoot "printer_driver_list.csv"
$currentHost = [Environment]::GetEnvironmentVariable("SELECTED_NEW_PCNAME")

if (Test-Path $driverCsvPath) {
    $csvDrivers = Import-ModuleCsv -Path $driverCsvPath -FilterEnabled `
        -RequiredColumns @("Enabled", "TargetHost", "DriverName")

    if ($null -eq $csvDrivers) {
        Show-Warning "Failed to load printer_driver_list.csv (continuing with hostlist only)"
    }
    else {
        $csvMatched = 0
        foreach ($row in @($csvDrivers)) {
            $targetHost = if ($null -ne $row.TargetHost) { $row.TargetHost.Trim() } else { "" }
            $isAllHosts = [string]::IsNullOrEmpty($targetHost)
            $isMatch = $isAllHosts -or ($targetHost -ieq $currentHost)

            if (-not $isMatch) { continue }

            $name = $row.DriverName
            if (-not [string]::IsNullOrEmpty($name) -and $name -notin $autoDriverNames) {
                $autoDriverNames += $name
                $csvMatched++
            }
        }
        if ($csvMatched -gt 0) {
            Show-Info "printer_driver_list.csv: $csvMatched new driver(s) added"
        }
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

    # ========================================
    # Step 5.5: Post-Apply Verification
    # ========================================
    # Read back the printer driver store and confirm each matched driver is registered.
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Post-Apply Verification" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $verifyPass = 0
    $verifyFail = 0

    foreach ($m in $matchedDrivers) {
        $actual = Get-PrinterDriver -Name $m.DriverName -ErrorAction SilentlyContinue
        if ($null -ne $actual) {
            Write-Host "  [VERIFIED] $($m.DriverName)" -ForegroundColor Green
            $verifyPass++
        }
        else {
            Write-Host "  [VERIFY FAILED] $($m.DriverName) - not found in driver store" -ForegroundColor Red
            $verifyFail++
        }
    }

    Write-Host ""
    $verified = ($verifyFail -eq 0)

    return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
        -Title "Installation Results" -Verified $verified)
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
