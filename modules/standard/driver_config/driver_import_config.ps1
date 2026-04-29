# ========================================
# Driver Import Script
# ========================================
# Import (install) drivers stored under driver/{model-name}/ via pnputil.
#
# [NOTES]
# - Requires administrator privileges
# - Uses pnputil.exe /add-driver
# - Exit code 3010 means "reboot required" and is treated as success
# ========================================

Write-Host ""
Show-Separator
Write-Host "Driver Import" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "driver.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Id", "model")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load driver.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: Prerequisite check (early return)
# ========================================
$driverDir = Join-Path $PSScriptRoot "driver"
if (-not (Test-Path $driverDir)) {
    Show-Error "Driver directory not found: $driverDir"
    Show-Error "Run Driver Export first to create driver backups."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Driver directory not found")
}


# ========================================
# Step 3: Dry-run summary before execution
# ========================================
# Resolve the system model name dynamically
$systemModel = ""
try {
    $systemModel = Get-CimInstance -ClassName Win32_ComputerSystem |
        Select-Object -ExpandProperty Model
}
catch {
    Show-Warning "Failed to get system model: $_"
}

# Sanitize a model name into a path-safe form
function Get-SafeModelName {
    param([string]$RawName)
    $safeName = $RawName -replace '\s', '_'
    $safeName = $safeName -replace '[\\/:*?"<>|]', ''
    $safeName = $safeName.Trim('_').Trim('.')
    if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0, 80) }
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "Unknown_Model" }
    return $safeName
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Import Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$hasValidTarget = $false

foreach ($item in $enabledItems) {
    # Resolve the model name (CSV value wins, otherwise auto-detect)
    $modelName = if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        $item.model
    }
    else {
        Get-SafeModelName -RawName $systemModel
    }
    $sourcePath = Join-Path $driverDir $modelName

    if (-not (Test-Path $sourcePath)) {
        Write-Host "  [NOT FOUND] Id=$($item.Id) : $modelName" -ForegroundColor DarkGray
        Write-Host "    Path: $sourcePath" -ForegroundColor DarkGray
    }
    else {
        $infCount = @(Get-ChildItem -Path $sourcePath -Filter "*.inf" -Recurse -File).Count
        if ($infCount -eq 0) {
            Write-Host "  [EMPTY] Id=$($item.Id) : $modelName" -ForegroundColor DarkGray
            Write-Host "    Path: $sourcePath (no .inf files)" -ForegroundColor DarkGray
        }
        else {
            Write-Host "  [IMPORT] Id=$($item.Id) : $modelName" -ForegroundColor Yellow
            Write-Host "    Path: $sourcePath ($infCount .inf files)" -ForegroundColor DarkGray
            $hasValidTarget = $true
        }
    }

    # Show where the model name came from
    if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        Write-Host "    Source: CSV specified" -ForegroundColor DarkGray
    }
    else {
        Write-Host "    Source: Auto-detected ($systemModel)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (-not $hasValidTarget) {
    Show-Skip "No valid driver folders found to import"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No valid driver folders found")
}


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Import drivers from the above paths?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Apply-settings loop
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $enabledItems) {
    # Resolve the model name (CSV value wins, otherwise auto-detect)
    $modelName = if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        $item.model
    }
    else {
        Get-SafeModelName -RawName $systemModel
    }
    $sourcePath = Join-Path $driverDir $modelName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Importing: Id=$($item.Id) - $modelName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Skip when the folder does not exist
    if (-not (Test-Path $sourcePath)) {
        Show-Skip "Folder not found: $sourcePath"
        Write-Host ""
        $skipCount++
        continue
    }

    # Skip when there are no .inf files in the folder
    $infCount = @(Get-ChildItem -Path $sourcePath -Filter "*.inf" -Recurse -File).Count
    if ($infCount -eq 0) {
        Show-Skip "No .inf files in: $sourcePath"
        Write-Host ""
        $skipCount++
        continue
    }

    try {
        Show-Info "Running pnputil /add-driver ($infCount .inf files)..."
        $infPattern = Join-Path $sourcePath "*.inf"
        $pnputilResult = & pnputil.exe /add-driver $infPattern /subdirs /install 2>&1
        $exitCode = $LASTEXITCODE

        # Exit code interpretation:
        #   0    = success
        #   259  = already up to date / nothing to add (ERROR_NO_MORE_ITEMS)
        #   3010 = reboot required (treated as success)
        #   else = error
        if ($exitCode -eq 0) {
            Show-Success "Imported: $modelName"
            $successCount++
        }
        elseif ($exitCode -eq 259) {
            Show-Skip "Already up to date: $modelName (exit code 259)"
            $skipCount++
        }
        elseif ($exitCode -eq 3010) {
            Show-Success "Imported: $modelName (restart required)"
            Show-Warning "A system restart is required to complete driver installation."
            $successCount++
        }
        else {
            Show-Error "pnputil exited with code $exitCode"
            $pnputilOutput = $pnputilResult | Out-String
            if (-not [string]::IsNullOrWhiteSpace($pnputilOutput)) {
                Write-Host $pnputilOutput -ForegroundColor DarkGray
            }
            $failCount++
        }
    }
    catch {
        Show-Error "Failed: Id=$($item.Id) - $modelName : $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Driver Import Results")
