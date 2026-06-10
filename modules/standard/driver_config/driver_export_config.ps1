# ========================================
# Driver Export Script
# ========================================
# Export the PC's third-party drivers into driver/{model-name}/.
#
# [NOTES]
# - Requires administrator privileges
# - dism.exe /export-driver only exports third-party drivers
# - Existing folders are cleared and re-exported
# ========================================

Write-Host ""
Show-Separator
Write-Host "Driver Export" -ForegroundColor Cyan
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
    try {
        New-Item -Path $driverDir -ItemType Directory -Force | Out-Null
        Show-Info "Created driver directory: $driverDir"
    }
    catch {
        Show-Error "Failed to create driver directory: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create driver directory")
    }
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
Write-Host "Export Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $enabledItems) {
    # Resolve the model name (CSV value wins, otherwise auto-detect)
    $modelName = if (-not [string]::IsNullOrWhiteSpace($item.model)) {
        $item.model
    }
    else {
        Get-SafeModelName -RawName $systemModel
    }

    # Destructive path guard (CLAUDE.md section 8): the CSV model value
    # bypasses Get-SafeModelName, so validate it as a single path
    # component. Reject (don't transform) - legitimate values stay
    # byte-identical, keeping folder naming consistent with driver_import.
    if (-not [string]::IsNullOrWhiteSpace($item.model) -and
        -not (Test-FabriqSafePathComponent -Value $item.model)) {
        Write-Host "  [INVALID] Id=$($item.Id) : '$($item.model)'" -ForegroundColor Red
        Write-Host "    Reason: model is not a safe path component - will be recorded as Fail" -ForegroundColor Red
        Write-Host ""
        continue
    }

    $destPath = Join-Path $driverDir $modelName

    # Switch the display based on whether the folder exists
    if (Test-Path $destPath) {
        Write-Host "  [OVERWRITE] Id=$($item.Id) : $modelName" -ForegroundColor Red
    }
    else {
        Write-Host "  [EXPORT] Id=$($item.Id) : $modelName" -ForegroundColor Yellow
    }
    Write-Host "    Path: $destPath" -ForegroundColor DarkGray

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


# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Export drivers to the above paths?"
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

    # Destructive path guard (CLAUDE.md section 8): same check as the
    # preview - record Fail and never touch the filesystem.
    if (-not [string]::IsNullOrWhiteSpace($item.model) -and
        -not (Test-FabriqSafePathComponent -Value $item.model)) {
        Show-Error "Invalid model name in CSV: '$($item.model)' (not a safe path component)"
        $failCount++
        continue
    }

    $destPath = Join-Path $driverDir $modelName

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Exporting: Id=$($item.Id) - $modelName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    try {
        # Containment assert: never delete anything outside the module's
        # driver directory (belt-and-suspenders behind the guard above).
        $dirFull  = [System.IO.Path]::GetFullPath($driverDir).TrimEnd('\').ToLowerInvariant()
        $destFull = [System.IO.Path]::GetFullPath($destPath).TrimEnd('\').ToLowerInvariant()
        if (-not $destFull.StartsWith($dirFull + '\')) {
            throw "Destination escapes driver directory: $destPath"
        }

        # Clear any existing folder
        if (Test-Path $destPath) {
            Show-Info "Clearing existing folder: $destPath"
            Remove-Item -Path $destPath -Recurse -Force
        }

        # Create the folder
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null

        # Export-WindowsDriver cmdlet fails with SafeHandle null on Server 2022;
        # dism.exe wraps the same DismApi and produces identical output
        Show-Info "Running dism.exe /online /export-driver..."
        $dismOutput = & dism.exe /online /export-driver /destination:"$destPath" 2>&1
        $dismExitCode = $LASTEXITCODE
        if ($dismExitCode -ne 0) {
            $lastMeaningful = $dismOutput |
                Where-Object { $_ -and ($_.ToString().Trim().Length -gt 0) } |
                Select-Object -Last 1
            throw "dism.exe /export-driver exited with code $dismExitCode : $lastMeaningful"
        }

        # Verify the export result
        $infCount = @(Get-ChildItem -Path $destPath -Filter "*.inf" -Recurse -File).Count
        Show-Success "Exported to: $destPath ($infCount .inf files)"
        $successCount++
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
    -Title "Driver Export Results")
