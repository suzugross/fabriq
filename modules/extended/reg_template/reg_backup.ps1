# ========================================
# Registry Backup (Template)
# ========================================
# Exports registry keys defined in reg_list.csv to .reg files.
# Copy this directory and edit reg_list.csv to create
# custom registry backup modules.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Registry Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# --- CSV Loading ---
$csvPath = Join-Path $PSScriptRoot "reg_list.csv"

$items = Import-ModuleCsv -Path $csvPath -FilterEnabled
if ($null -eq $items) { return (New-ModuleResult -Status "Error" -Message "Failed to load reg_list.csv") }
if ($items.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled entries") }

# --- Backup directory ---
$backupDir = Join-Path $PSScriptRoot "backup"
if (-not (Test-Path $backupDir)) {
    try {
        $null = New-Item -ItemType Directory -Path $backupDir -Force
    }
    catch {
        Show-Error "Failed to create backup directory: $_"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Failed to create backup dir")
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# --- Display target list ---
Show-Info "Backup targets: $($items.Count) registry keys"
Write-Host ""

$index = 0
foreach ($item in $items) {
    $index++
    $keyExists = $false
    try {
        $checkPath = $item.RegistryPath -replace '^HKEY_LOCAL_MACHINE', 'HKLM:' -replace '^HKEY_CURRENT_USER', 'HKCU:' -replace '^HKEY_CLASSES_ROOT', 'HKCR:' -replace '^HKEY_USERS', 'HKU:'
        $keyExists = Test-Path $checkPath -ErrorAction SilentlyContinue
    }
    catch { }

    $marker = if ($keyExists) { "[OK]" } else { "[--]" }
    $markerColor = if ($keyExists) { "Green" } else { "Yellow" }

    Write-Host "  [$index] $($item.RegistryPath)  $marker" -ForegroundColor $markerColor
    if ($item.Description) {
        Write-Host "      $($item.Description)" -ForegroundColor Gray
    }
}
Write-Host ""

# --- Confirmation ---
$cancelResult = Confirm-ModuleExecution -Message "Export the above registry keys?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# --- Execute backup ---
$successCount = 0
$failCount = 0
$total = $items.Count
$current = 0

foreach ($item in $items) {
    $current++
    $regPath = $item.RegistryPath.Trim()

    # Generate filename from registry path (sanitize)
    $safeName = $regPath -replace '[\\/:*?"<>|]', '_'
    if ($safeName.Length -gt 80) { $safeName = $safeName.Substring(0, 80) }
    $exportFile = Join-Path $backupDir "${timestamp}_${safeName}.reg"

    Write-Host "[$current/$total] $regPath" -ForegroundColor Cyan

    try {
        $process = Start-Process reg.exe -ArgumentList "export `"$regPath`" `"$exportFile`" /y" -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -eq 0) {
            Show-Success "Exported: $([System.IO.Path]::GetFileName($exportFile))"
            $successCount++
        }
        else {
            Show-Error "reg.exe exit code: $($process.ExitCode)"
            $failCount++
        }
    }
    catch {
        Show-Error "$($_.Exception.Message)"
        $failCount++
    }

    Write-Host ""
}

# --- Summary ---
Show-Info "Location: $backupDir"
Write-Host ""
return (New-BatchResult -Success $successCount -Fail $failCount -Title "Registry Backup Results")
