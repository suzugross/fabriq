# ========================================
# Local Group Policy Backup Script
# ========================================
# [PURPOSE]
# Reads the local Group Policy registry policy files
# (%SystemRoot%\System32\GroupPolicy\{Machine,User}\Registry.pol) and
# writes them out as a gpo_list-format CSV so settings made on a
# reference PC (gpedit.msc) can be replayed by gpo_config.ps1.
# The raw Registry.pol files and gpt.ini are copied alongside.
#
# [NOTES]
# - Read-only against the system; output goes to backup\<timestamp>\.
# - Entries using policy directives other than **del. / **delvals.
#   (e.g. **DeleteKeys) are written with Enabled=0 and flagged.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Local Group Policy Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""

. (Join-Path $PSScriptRoot "lib\PolFile.ps1")

# ========================================
# Step 1: Discover policy files (no CSV input for this script)
# ========================================
$sources = @()
foreach ($scopeName in @('Machine', 'User')) {
    $polPath = Get-PolFilePath -Scope $scopeName
    $exists = Test-Path -LiteralPath $polPath
    $marker = if ($exists) { "[OK]" } else { "[--]" }
    $markerColor = if ($exists) { "Green" } else { "Yellow" }
    Write-Host "  $($scopeName.PadRight(8)) $polPath  $marker" -ForegroundColor $markerColor
    if ($exists) { $sources += [PSCustomObject]@{ Scope = $scopeName; Path = $polPath } }
}
Write-Host ""

if ($sources.Count -eq 0) {
    Show-Info "No local group policy registry files found"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No local Registry.pol found")
}

# ========================================
# Step 2: Prerequisite check - backup directory
# ========================================
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = Join-Path (Join-Path $PSScriptRoot "backup") $timestamp
try {
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
}
catch {
    Show-Error "Failed to create backup directory: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to create backup dir")
}

# ========================================
# Step 3: Dry-run summary
# ========================================
$parsed = @()
foreach ($src in $sources) {
    try {
        $entries = @(Read-PolFile -Path $src.Path)
        $parsed += [PSCustomObject]@{ Scope = $src.Scope; Path = $src.Path; Entries = $entries; Error = $null }
        Write-Host "  [PARSE] $($src.Scope): $(@($entries).Count) entries" -ForegroundColor Yellow
    }
    catch {
        $parsed += [PSCustomObject]@{ Scope = $src.Scope; Path = $src.Path; Entries = $null; Error = $_.Exception.Message }
        Write-Host "  [ERROR] $($src.Scope): $($_.Exception.Message)" -ForegroundColor Red
    }
}
Write-Host ""
Show-Info "Output: $backupDir"
Write-Host ""

# ========================================
# Step 4: User confirmation
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Export the local group policy above to CSV?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Export
# ========================================
$successCount = 0
$failCount    = 0
$csvRows      = New-Object System.Collections.Generic.List[object]
$adminId      = 0
$unsupported  = 0

foreach ($p in $parsed) {
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $($p.Scope) policy" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    if ($null -ne $p.Error) {
        Show-Error "Skipping unreadable file: $($p.Error)"
        $failCount++
        Write-Host ""
        continue
    }

    try {
        Copy-Item -LiteralPath $p.Path -Destination (Join-Path $backupDir "$($p.Scope)_Registry.pol") -Force

        foreach ($e in @($p.Entries)) {
            $adminId++
            $action = Get-PolEntryAction -Entry $e
            $enabled = '1'
            $valueName = [string]$e.ValueName
            $typeName = ''
            $value = ''

            switch ($action) {
                'Set' {
                    $typeName = Get-PolTypeName -Code $e.Type
                    $value = ConvertFrom-PolData -Type $e.Type -Data $e.Data
                }
                'Delete'          { $valueName = $valueName -replace '^\*\*del\.', '' }
                'DeleteAllValues' { $valueName = '' }
                'CreateKey'       { $valueName = '' }
                'Unsupported' {
                    $enabled = '0'
                    $unsupported++
                    Show-Warning "Unsupported policy directive '$valueName' under $($e.Key) - exported with Enabled=0"
                }
            }

            $title = "$($p.Scope) $($e.Key)"
            if ($valueName) { $title += "\$valueName" }
            $title += " [$action]"

            $csvRows.Add([PSCustomObject]@{
                Enabled      = $enabled
                AdminID      = $adminId
                SettingTitle = $title
                Scope        = $p.Scope
                KeyPath      = $e.Key
                ValueName    = $valueName
                Action       = $action
                Type         = $typeName
                Value        = $value
                PolicyRef    = ''
                Segment      = ''
            })
        }

        Show-Success "$($p.Scope): $(@($p.Entries).Count) entries exported"
        $successCount++
    }
    catch {
        Show-Error "$($p.Scope): $($_.Exception.Message)"
        $failCount++
    }
    Write-Host ""
}

$gptIni = Get-GptIniPath
if (Test-Path -LiteralPath $gptIni) {
    Copy-Item -LiteralPath $gptIni -Destination (Join-Path $backupDir "gpt.ini") -Force
}

$csvPath = Join-Path $backupDir "gpo_list_backup.csv"
$csvWritten = $false
try {
    $header = 'Enabled,AdminID,SettingTitle,Scope,KeyPath,ValueName,Action,Type,Value,PolicyRef,Segment'
    $lines = @($header)
    if ($csvRows.Count -gt 0) {
        $lines += @($csvRows | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1)
    }
    # UTF-8 with BOM + CRLF (fabriq CSV convention)
    [IO.File]::WriteAllText($csvPath, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    $csvWritten = $true
    Show-Success "CSV written: $csvPath ($($csvRows.Count) rows)"
}
catch {
    Show-Error "Failed to write CSV: $($_.Exception.Message)"
    $failCount++
}
Write-Host ""

# ========================================
# Step 5.5: Post-Apply Verification - CSV round-trip row count
# ========================================
$verified = $null
if ($csvWritten) {
    try {
        $check = @(Import-Csv -LiteralPath $csvPath -Encoding UTF8)
        $verified = ($check.Count -eq $csvRows.Count)
        if ($verified) {
            Write-Host "  [VERIFIED] CSV re-read: $($check.Count) rows" -ForegroundColor Green
        }
        else {
            Write-Host "  [VERIFY FAILED] CSV re-read returned $($check.Count) rows, expected $($csvRows.Count)" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "  [VERIFY FAILED] CSV re-read failed: $($_.Exception.Message)" -ForegroundColor Red
        $verified = $false
    }
    Write-Host ""
}

if ($unsupported -gt 0) {
    Show-Warning "$unsupported entries use unsupported directives and were exported with Enabled=0"
    Write-Host ""
}

# ========================================
# Step 6: Aggregate and return result
# ========================================
return (New-BatchResult -Success $successCount -Fail $failCount `
    -Title "Local Group Policy Backup Results" -Verified $verified)
