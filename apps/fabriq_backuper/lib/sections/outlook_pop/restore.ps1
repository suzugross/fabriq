# ============================================================
# FabriqBackUper Section: outlook_pop / restore (Phase 2.9.0 Phase A)
#
# STUB. Phase A ships backup-only. Phase B will:
#   - Read sections/outlook_pop/manifest.json
#   - Kill any running OUTLOOK.EXE
#   - For each POP account, generate a temporary PRF file
#     (Microsoft canonical 7-section format with
#     [Service2 = Internet E-mail] + [AccountN] blocks)
#   - Invoke `outlook.exe /importprf <prf>` for each
#   - Delete the PRF file (no password is embedded anyway)
#   - Operator re-enters passwords on first send/receive in Outlook
#
# Returning Status='Skipped' lets the engine + UI display a clean
# placeholder row rather than treating it as a failure.
# ============================================================

param(
    [Parameter(Mandatory = $true)][string]$BackuperRoot,
    [Parameter(Mandatory = $true)][string]$FabriqRoot,
    [Parameter(Mandatory = $true)][string]$OldPcName,
    [Parameter(Mandatory = $true)][string]$AggregateBackupDir,
    [hashtable]$SectionParams = @{}
)

$sw = [System.Diagnostics.Stopwatch]::StartNew()

$sectionDir = Join-Path $AggregateBackupDir 'sections\outlook_pop'
$manifestPath = Join-Path $sectionDir 'manifest.json'
$manifestPresent = Test-Path $manifestPath

$msg = if ($manifestPresent) {
    'outlook_pop restore is not yet implemented (Phase B). Manifest is present; ' +
    'use it to manually craft a PRF and run: outlook.exe /importprf <prf>'
} else {
    'outlook_pop restore stub: no manifest found at expected path.'
}
Show-Info $msg

$sw.Stop()

return [PSCustomObject]@{
    Status               = 'Skipped'
    ElapsedMs            = [int]$sw.ElapsedMilliseconds
    Summary              = [ordered]@{
        note            = 'restore not yet implemented (Phase B)'
        manifestPresent = $manifestPresent
    }
    Warnings             = @()
    ExternalOutputDir    = $null
    ExternalManifestPath = $null
}
