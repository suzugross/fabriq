# ========================================
# Registry Collection App  -  Phase 1
# ========================================
# Browse a catalog of registry settings and
# append selected entries to reg_config CSVs.
# UI: console + Out-GridView (no WinForms)
# Phase 2 will add a full WinForms GUI.
# ========================================

# ========================================
# Fabriq Root & Common Import
# ========================================
$script:FabriqRoot = Split-Path (Split-Path $PSScriptRoot -Parent)
$commonPath = Join-Path $script:FabriqRoot "kernel\common.ps1"
if (Test-Path $commonPath) {
    . $commonPath
}
else {
    Write-Host "[ERROR] kernel/common.ps1 not found: $commonPath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# ========================================
# Constants
# ========================================
$script:CatalogPath = Join-Path $PSScriptRoot "catalog.csv"
$script:DocsDir     = Join-Path $PSScriptRoot "docs"
$script:HklmCsvPath = Join-Path $script:FabriqRoot "modules\standard\reg_config\reg_hklm_list.csv"
$script:HkcuCsvPath = Join-Path $script:FabriqRoot "modules\standard\reg_config\reg_hkcu_list.csv"

$script:CatalogRequiredColumns = @(
    "Category", "Title", "Hive", "KeyPath", "KeyName", "Type", "Value"
)

# ========================================
# Function: Read-Catalog
# ========================================
# Loads catalog.csv and returns all entries.
# Optionally filters by Category or Hive.
# Returns $null on load failure.
# ========================================
function Read-Catalog {
    param(
        [string]$CategoryFilter = "",
        [string]$HiveFilter     = ""
    )

    if (-not (Test-Path $script:CatalogPath)) {
        Show-Error "Catalog not found: $script:CatalogPath"
        return $null
    }

    $entries = Import-CsvSafe -Path $script:CatalogPath
    if ($null -eq $entries) {
        Show-Error "Failed to load catalog.csv"
        return $null
    }

    if (-not (Test-CsvColumns -CsvData $entries -RequiredColumns $script:CatalogRequiredColumns)) {
        Show-Error "catalog.csv is missing required columns: $($script:CatalogRequiredColumns -join ', ')"
        return $null
    }

    if ($CategoryFilter -ne "") {
        $entries = @($entries | Where-Object { $_.Category -eq $CategoryFilter })
    }
    if ($HiveFilter -ne "") {
        $entries = @($entries | Where-Object { $_.Hive -eq $HiveFilter })
    }

    # Add display-only column (shortened KeyPath) — preserved through Out-GridView -PassThru
    $entries = $entries | ForEach-Object {
        $_ | Add-Member -NotePropertyName "KeyPath_Short" `
            -NotePropertyValue (
                $_.KeyPath `
                    -replace "HKEY_LOCAL_MACHINE", "HKLM" `
                    -replace "HKEY_CURRENT_USER",  "HKCU"
            ) -PassThru -Force
    }

    return $entries
}

# ========================================
# Function: Show-CatalogDoc
# ========================================
# Displays the description file for an entry.
# Falls back gracefully when DocFile is absent.
# ========================================
function Show-CatalogDoc {
    param(
        [Parameter(Mandatory = $true)]
        $Entry
    )

    Write-Host ""
    Show-Separator
    Write-Host "  $($Entry.Title)" -ForegroundColor Cyan
    Write-Host "  [$($Entry.Hive)] $($Entry.KeyPath)" -ForegroundColor DarkGray
    Write-Host "  $($Entry.KeyName) = [$($Entry.Type)] $($Entry.Value)" -ForegroundColor White
    Show-Separator

    if ([string]::IsNullOrWhiteSpace($Entry.DocFile)) {
        Show-Info "No description file registered for this entry"
        return
    }

    $docPath = Join-Path $script:DocsDir $Entry.DocFile
    if (-not (Test-Path $docPath)) {
        Show-Warning "Description file not found: $($Entry.DocFile)"
        return
    }

    $content = Get-Content -Path $docPath -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($null -eq $content) {
        Show-Warning "Could not read description file: $($Entry.DocFile)"
        return
    }

    foreach ($line in $content) {
        Write-Host "  $line" -ForegroundColor Gray
    }
    Write-Host ""
}

# ========================================
# Function: Export-ToRegConfig
# ========================================
# Appends selected catalog entries to the
# appropriate reg_config CSV files.
# - Splits entries by Hive (HKLM / HKCU)
# - Auto-increments AdminID from existing max
# - Skips duplicates (same KeyPath + KeyName)
# Returns @{ Added = N; Skipped = N }
# ========================================
function Export-ToRegConfig {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Entries,
        [string]$HklmPath = $script:HklmCsvPath,
        [string]$HkcuPath = $script:HkcuCsvPath
    )

    $addedCount   = 0
    $skippedCount = 0

    $hklmEntries = @($Entries | Where-Object { $_.Hive -eq "HKLM" })
    $hkcuEntries = @($Entries | Where-Object { $_.Hive -eq "HKCU" })

    $groups = @(
        @{ Entries  = $hklmEntries; CsvPath = $HklmPath; HiveName = "HKLM" }
        @{ Entries  = $hkcuEntries; CsvPath = $HkcuPath; HiveName = "HKCU" }
    )

    foreach ($group in $groups) {
        if ($group.Entries.Count -eq 0) { continue }

        $csvPath  = $group.CsvPath
        $hiveName = $group.HiveName

        # Load existing rows for duplicate detection and AdminID calculation
        $existingKeys = @{}
        $maxAdminId   = 0

        if (Test-Path $csvPath) {
            $existingRows = @(Import-Csv -Path $csvPath -Encoding UTF8 -ErrorAction SilentlyContinue)
            foreach ($row in $existingRows) {
                $dupKey = "$($row.KeyPath)|$($row.KeyName)"
                $existingKeys[$dupKey] = $true
                $id = [int]($row.AdminID -as [int])
                if ($id -gt $maxAdminId) { $maxAdminId = $id }
            }
        }
        else {
            Show-Warning "[$hiveName] CSV not found — will create: $csvPath"
        }

        $nextId = $maxAdminId + 1

        foreach ($entry in $group.Entries) {
            $dupKey = "$($entry.KeyPath)|$($entry.KeyName)"

            if ($existingKeys.ContainsKey($dupKey)) {
                Show-Skip "[$hiveName] Already exists — skipped: $($entry.Title)"
                $skippedCount++
                continue
            }

            # Build output row matching reg_config CSV format exactly
            $newRow = [PSCustomObject][ordered]@{
                Enabled      = "1"
                AdminID      = "$nextId"
                SettingTitle = $entry.Title
                KeyPath      = $entry.KeyPath
                KeyName      = $entry.KeyName
                Type         = $entry.Type
                Value        = $entry.Value
            }

            try {
                $newRow | Export-Csv -Path $csvPath -Append -NoTypeInformation `
                    -Encoding UTF8 -ErrorAction Stop
                Show-Success "[$hiveName] Added (ID=$nextId): $($entry.Title)"
                $existingKeys[$dupKey] = $true
                $nextId++
                $addedCount++
            }
            catch {
                Show-Error "[$hiveName] Write failed: $($entry.Title) — $_"
            }
        }
    }

    return @{ Added = $addedCount; Skipped = $skippedCount }
}

# ========================================
# Main: Start-RegistryCollectionApp
# ========================================
function Start-RegistryCollectionApp {
    Clear-Host
    Write-Host ""
    Show-Separator
    Write-Host "  Registry Collection App" -ForegroundColor Cyan
    Write-Host "  reg_config CSV Builder  [Phase 1]" -ForegroundColor DarkGray
    Show-Separator
    Write-Host ""

    # ---- Step 1: Load catalog ----
    Show-Info "Loading catalog..."
    $catalog = Read-Catalog
    if ($null -eq $catalog -or @($catalog).Count -eq 0) {
        Show-Error "No entries found in catalog. Exiting."
        Read-Host "Press Enter to exit"
        return
    }
    Show-Success "Catalog loaded: $(@($catalog).Count) entries"
    Write-Host ""

    # ---- Step 2: Select entries via Out-GridView ----
    Show-Info "An entry browser will open."
    Write-Host "  - Use the filter box to search (Category, Title, Tags, etc.)" -ForegroundColor DarkGray
    Write-Host "  - Ctrl+Click or Shift+Click to select multiple entries" -ForegroundColor DarkGray
    Write-Host "  - Click [OK] to proceed with selected entries" -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press Enter to open browser"

    $selected = @($catalog | Out-GridView `
        -Title "Registry Collection — Select entries to add to reg_config CSV  [Ctrl+Click = multi-select]" `
        -PassThru)

    if ($selected.Count -eq 0) {
        Show-Info "No entries selected. Exiting."
        return
    }

    $selectedCount = $selected.Count
    Write-Host ""
    Show-Info "Selected: $selectedCount entries"
    Write-Host ""

    # ---- Step 3: Show doc for each selected entry ----
    foreach ($entry in $selected) {
        Show-CatalogDoc -Entry $entry
    }

    # ---- Step 4: Confirm ----
    $shouldExport = Confirm-Execution -Message "Add $selectedCount entries to reg_config CSV?"
    if (-not $shouldExport) {
        Show-Info "Cancelled. No changes made."
        return
    }

    # ---- Step 5: Export ----
    Write-Host ""
    Show-Separator
    $result = Export-ToRegConfig -Entries $selected
    Show-Separator
    Write-Host ""

    # ---- Step 6: Result summary ----
    if ($result.Added -gt 0) {
        Show-Success "Done — Added: $($result.Added)  Skipped: $($result.Skipped)"
    }
    else {
        Show-Info "Done — Added: $($result.Added)  Skipped: $($result.Skipped)"
    }
    Write-Host ""
    Write-Host "  Output files:" -ForegroundColor DarkGray
    Write-Host "    HKLM -> $script:HklmCsvPath" -ForegroundColor DarkGray
    Write-Host "    HKCU -> $script:HkcuCsvPath" -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press Enter to exit"
}

# ========================================
# Entry Point
# ========================================
Start-RegistryCollectionApp
