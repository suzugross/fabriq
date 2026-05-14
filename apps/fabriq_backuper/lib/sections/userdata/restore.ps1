# ============================================================
# FabriqBackUper Section: userdata / restore (Phase 2.3, internalized)
#
# Reads $AggregateBackupDir/sections/userdata/manifest.json
# (fabriq-userdata-backup schemaVersion=1) and replays each entry
# back to its resolvedPath via robocopy.
#
# SectionParams (hashtable, all optional):
#   IncludeEntries  : array of SourcePath strings to restore
#                     (null/empty = all entries in manifest, except 'Skipped')
# ============================================================

param(
    [Parameter(Mandatory = $true)][string]$BackuperRoot,
    [Parameter(Mandatory = $true)][string]$FabriqRoot,
    [Parameter(Mandatory = $true)][string]$OldPcName,
    [Parameter(Mandatory = $true)][string]$AggregateBackupDir,
    [hashtable]$SectionParams = @{}
)

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$warnings = @()

$includeEntries = $null
if ($SectionParams.ContainsKey('IncludeEntries') -and `
    $null -ne $SectionParams['IncludeEntries'] -and `
    @($SectionParams['IncludeEntries']).Count -gt 0) {
    $includeEntries = @($SectionParams['IncludeEntries'])
}

# ----------------------------------------------------------
# Prereq + manifest validation
# ----------------------------------------------------------
if (-not (Test-AdminPrivilege)) {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @('Administrator privileges required')
    }
}
$robocopyExe = Get-Command robocopy.exe -ErrorAction SilentlyContinue
if ($null -eq $robocopyExe) {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @('robocopy.exe not found')
    }
}

$sectionDir = Join-Path $AggregateBackupDir 'sections\userdata'
$manifestPath = Join-Path $sectionDir 'manifest.json'
if (-not (Test-Path $manifestPath)) {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @("manifest.json not found at: $manifestPath")
    }
}

$manifest = $null
try { $manifest = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json } catch {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @("manifest.json parse error: $($_.Exception.Message)")
    }
}
if ($manifest.manifestType -ne 'fabriq-userdata-backup') {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @("Unexpected manifestType: $($manifest.manifestType)")
    }
}
if ([int]$manifest.schemaVersion -ne 1) {
    return [PSCustomObject]@{
        Status = 'Failed'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{}; Warnings = @("Unsupported schemaVersion: $($manifest.schemaVersion)")
    }
}

# ----------------------------------------------------------
# Build restore plan
# ----------------------------------------------------------
$allEntries = @($manifest.items.entries)
$plannedEntries = @($allEntries | Where-Object {
    $_.status -ne 'Skipped' -and -not [string]::IsNullOrWhiteSpace($_.backupSubpath)
})
if ($null -ne $includeEntries) {
    $plannedEntries = @($plannedEntries | Where-Object { $_.sourcePath -in $includeEntries })
}

if ($plannedEntries.Count -eq 0) {
    return [PSCustomObject]@{
        Status = 'Skipped'; ElapsedMs = [int]$sw.ElapsedMilliseconds
        Summary = [ordered]@{ note = 'no restorable entries' }
        Warnings = @($warnings)
    }
}

Show-Info "Restoring $($plannedEntries.Count) entry(ies)"

# ----------------------------------------------------------
# Execute
# ----------------------------------------------------------
$successCount = 0; $skipCount = 0; $failCount = 0

foreach ($pe in $plannedEntries) {
    Show-Info "[$($pe.id)] $($pe.resolvedPath)"
    $srcDataDir = Join-Path $sectionDir $pe.backupSubpath
    if (-not (Test-Path $srcDataDir)) {
        Show-Error "  Backup data folder missing: $srcDataDir"
        $warnings += "Missing data folder for entry $($pe.id)"
        $failCount++
        continue
    }

    # Phase 2.4: prefer re-expanding sourcePath against the CURRENT target
    # process context, so cross-user migration (backup-user != restore-user)
    # resolves %USERPROFILE% etc. to the correct directory. Fall back to
    # the manifest's resolvedPath when sourcePath has no env vars.
    $targetPath = if ($pe.sourcePath -match '%\w+%') {
        [Environment]::ExpandEnvironmentVariables($pe.sourcePath)
    } else {
        $pe.sourcePath
    }
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        $targetPath = $pe.resolvedPath  # final fallback
    }
    Show-Info "  target: $targetPath  (source: $($pe.sourcePath))"

    $isDir      = [bool]$pe.isDirectory
    $onConflict = if ([string]::IsNullOrWhiteSpace($pe.onConflict)) { 'skip' } else { $pe.onConflict.ToLower() }
    $includeAcl = [bool]$pe.includeAcl
    $proceed = $true

    if ($isDir) {
        if (Test-Path -LiteralPath $targetPath -PathType Container) {
            switch ($onConflict) {
                'skip' { Show-Skip "  exists (skip)"; $skipCount++; $proceed = $false }
                'overwrite' { Show-Info "  exists (overwrite)" }
                'rename' {
                    $renamed = "$targetPath.bak_$(Get-Date -Format yyyyMMdd_HHmmss)"
                    try {
                        Rename-Item -LiteralPath $targetPath -NewName (Split-Path $renamed -Leaf) -ErrorAction Stop
                        Show-Info "  renamed existing -> $(Split-Path $renamed -Leaf)"
                    } catch {
                        $warnings += "Rename failed for $targetPath"; $failCount++; $proceed = $false
                    }
                }
                default { $warnings += "Unknown OnConflict for $($pe.id)"; $failCount++; $proceed = $false }
            }
        } elseif (Test-Path -LiteralPath $targetPath -PathType Leaf) {
            $warnings += "Type mismatch (file vs dir) for $($pe.id)"; $failCount++; $proceed = $false
        }
    } else {
        if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
            switch ($onConflict) {
                'skip' { Show-Skip "  exists (skip)"; $skipCount++; $proceed = $false }
                'overwrite' { Show-Info "  exists (overwrite)" }
                'rename' {
                    $renamed = "$targetPath.bak_$(Get-Date -Format yyyyMMdd_HHmmss)"
                    try {
                        Rename-Item -LiteralPath $targetPath -NewName (Split-Path $renamed -Leaf) -ErrorAction Stop
                        Show-Info "  renamed existing -> $(Split-Path $renamed -Leaf)"
                    } catch {
                        $warnings += "Rename failed for $targetPath"; $failCount++; $proceed = $false
                    }
                }
                default { $warnings += "Unknown OnConflict for $($pe.id)"; $failCount++; $proceed = $false }
            }
        }
    }

    if (-not $proceed) { continue }

    if ($isDir) {
        if (-not (Test-Path -LiteralPath $targetPath)) {
            try { $null = New-Item -ItemType Directory -Path $targetPath -Force -ErrorAction Stop }
            catch { $warnings += "Target dir create failed: $targetPath"; $failCount++; continue }
        }
    } else {
        $parent = Split-Path -Path $targetPath -Parent
        if (-not (Test-Path -LiteralPath $parent)) {
            try { $null = New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop }
            catch { $warnings += "Target parent create failed: $parent"; $failCount++; continue }
        }
    }

    $copyFlag = if ($includeAcl) { '/COPYALL' } else { '/COPY:DAT' }
    if ($isDir) {
        $rcArgs = @($srcDataDir, $targetPath, '/E', $copyFlag, '/B', '/R:1', '/W:1', '/NP')
    } else {
        $fileName = Split-Path -Path $targetPath -Leaf
        $targetDir = Split-Path -Path $targetPath -Parent
        $rcArgs = @($srcDataDir, $targetDir, $fileName, $copyFlag, '/B', '/R:1', '/W:1', '/NP', '/NDL', '/NS', '/NC')
    }
    & robocopy.exe @rcArgs 2>&1 | Out-Null
    $rcExit = $LASTEXITCODE
    if ($rcExit -ge 8) {
        Show-Error "  robocopy fail (exit $rcExit)"
        $warnings += "robocopy exit $rcExit for $($pe.id)"; $failCount++
    } elseif ($rcExit -ge 4) {
        Show-Warning "  robocopy mismatch (exit $rcExit)"
        $warnings += "robocopy mismatch (exit $rcExit) for $($pe.id)"; $successCount++
    } else {
        Show-Success "  restored (exit $rcExit)"
        $successCount++
    }
}

$sw.Stop()
$status = if ($failCount -gt 0 -and $successCount -eq 0) { 'Failed' }
          elseif ($failCount -gt 0) { 'Partial' }
          else { 'Success' }

return [PSCustomObject]@{
    Status   = $status
    ElapsedMs = [int]$sw.ElapsedMilliseconds
    Summary  = [ordered]@{
        entrySuccess = $successCount
        entrySkip    = $skipCount
        entryFail    = $failCount
    }
    Warnings = $warnings
}
