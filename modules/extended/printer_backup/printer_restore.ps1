# ========================================
# Printer Restore
# ========================================
# [PURPOSE]
# Replay a backup produced by modules/extended/printer_backup onto
# this PC. Reads manifest.json (schemaVersion=1, fabriq-printer-backup)
# and restores drivers, ports, printers, print settings, and the
# default printer in dependency order.
#
# [NOTES]
# - Requires administrator privileges.
# - Restore source: ./backup/<PCName>/<timestamp>/ (same module dir,
#   produced by the companion printer_backup.ps1).
# - Shares printer_backup_config.csv with printer_backup.ps1.
# - osArch mismatch is a hard fail (e.g. amd64 backup -> arm64 target).
# - osVersion mismatch is gated by StrictOsVersion column
#   (default 0 = warn but proceed, accommodates Win10 -> Win11 amd64).
# - RDP-redirected printers in the backup (driverName=
#   'Remote Desktop Easy Print' or portName=TS\d+) are skipped.
# - Inbox / Microsoft-supplied drivers re-use the Windows in-box
#   driver instead of re-installing from payload when
#   ReuseInboxDrivers=1.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Printer Restore" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load Config CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "printer_backup_config.csv"

$cfgItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @(
        "Enabled",
        "SourcePcName",
        "BackupTimestamp",
        "StrictOsVersion",
        "ReuseInboxDrivers",
        "OnConflict",
        "RestoreDefaultPrinter",
        "SkipVirtualPrinters",
        "RestoreHardwareConfig"
    )

if ($null -eq $cfgItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load printer_backup_config.csv")
}
if (@($cfgItems).Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "Restore disabled in config CSV")
}

$cfg = @($cfgItems)[0]
$cfgSourcePcName        = if ([string]::IsNullOrWhiteSpace($cfg.SourcePcName))    { $null } else { $cfg.SourcePcName.Trim() }
$cfgBackupTimestamp     = if ([string]::IsNullOrWhiteSpace($cfg.BackupTimestamp)) { $null } else { $cfg.BackupTimestamp.Trim() }
$strictOsVersion        = ($cfg.StrictOsVersion       -eq "1")
$reuseInboxDrivers      = ($cfg.ReuseInboxDrivers     -eq "1")
$restoreDefaultPrinter  = ($cfg.RestoreDefaultPrinter -eq "1")
$skipVirtualPrinters    = ($cfg.SkipVirtualPrinters   -eq "1")
$restoreHardwareConfig  = ($cfg.RestoreHardwareConfig -eq "1")
$onConflict             = if ([string]::IsNullOrWhiteSpace($cfg.OnConflict)) { "skip" } else { $cfg.OnConflict.Trim().ToLower() }
if ($onConflict -ne "skip" -and $onConflict -ne "replace") {
    Show-Error "Invalid OnConflict value: '$onConflict' (expected: skip / replace)"
    return (New-ModuleResult -Status "Error" -Message "Invalid OnConflict value")
}


# ========================================
# Step 2: Prerequisite Checks
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

try {
    $spooler = Get-Service -Name Spooler -ErrorAction Stop
    if ($spooler.Status -ne 'Running') {
        Show-Error "Print Spooler service is not running (Status: $($spooler.Status))"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Print Spooler not running")
    }
}
catch {
    Show-Error "Failed to check Print Spooler service: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Spooler service check failed")
}


# ========================================
# Step 3a: Locate Source Backup
# ========================================
# Resolution priority for the source PC name:
#   1. CSV SourcePcName (explicit override, for testing / cross-PC moves)
#   2. $env:SELECTED_OLD_PCNAME (hostlist context — the standard kitting
#      workflow path: operator selects a host, OldPCname is set, restore
#      auto-picks the matching backup subfolder)
#   3. Neither set -> hard error (we refuse the ambiguous "scan all PCs"
#      path because it makes it unclear which PC's settings were restored)
#
# Resolution priority for the timestamp:
#   1. CSV BackupTimestamp (explicit override)
#   2. Latest manifest.collectedAt under the resolved PcName
$backupRoot = Join-Path $PSScriptRoot "backup"
$backupRoot = [System.IO.Path]::GetFullPath($backupRoot)

if (-not (Test-Path $backupRoot)) {
    Show-Error "Backup root not found: $backupRoot"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Backup root not found")
}

Show-Info "Backup root: $backupRoot"

# Step 3a-i: Resolve target PC name
$envOldPcName = $env:SELECTED_OLD_PCNAME
$pcNameSource = $null
$resolvedPcName = $null
if (-not [string]::IsNullOrWhiteSpace($cfgSourcePcName)) {
    $resolvedPcName = $cfgSourcePcName
    $pcNameSource   = "CSV SourcePcName"
}
elseif (-not [string]::IsNullOrWhiteSpace($envOldPcName)) {
    $resolvedPcName = $envOldPcName.Trim()
    $pcNameSource   = "hostlist OldPCname"
}
else {
    Show-Error "No source PC name available."
    Show-Error "  Either select a host (so SELECTED_OLD_PCNAME is set) or"
    Show-Error "  set SourcePcName in printer_backup_config.csv."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No source PC name (host not selected, CSV SourcePcName empty)")
}

Show-Info "Source PC name : $resolvedPcName (from $pcNameSource)"

# Step 3a-ii: Locate matching <PcName> subfolder
$pcDir = Get-ChildItem -Path $backupRoot -Directory -ErrorAction SilentlyContinue |
         Where-Object { $_.Name -ieq $resolvedPcName } |
         Select-Object -First 1

if ($null -eq $pcDir) {
    Show-Error "No backup subfolder for PC '$resolvedPcName' under: $backupRoot"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No backup folder for source PC '$resolvedPcName'")
}

# Step 3a-iii: Enumerate timestamp folders containing manifest.json
$candidates = @()
foreach ($tsDir in @(Get-ChildItem -Path $pcDir.FullName -Directory -ErrorAction SilentlyContinue)) {
    if (-not [string]::IsNullOrWhiteSpace($cfgBackupTimestamp) -and ($tsDir.Name -ine $cfgBackupTimestamp)) { continue }
    $mfPath = Join-Path $tsDir.FullName "manifest.json"
    if (-not (Test-Path $mfPath)) { continue }
    $candidates += [PSCustomObject]@{
        PcName    = $pcDir.Name
        Timestamp = $tsDir.Name
        Path      = $tsDir.FullName
        Manifest  = $mfPath
    }
}

if ($candidates.Count -eq 0) {
    $filter = if (-not [string]::IsNullOrWhiteSpace($cfgBackupTimestamp)) {
        " (timestamp filter: '$cfgBackupTimestamp')"
    } else { "" }
    Show-Error "No backup folders with manifest.json under: $($pcDir.FullName)$filter"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No backups under PC '$resolvedPcName'")
}

# Step 3a-iv: Pick newest by manifest.collectedAt (fallback: folder mtime)
$chosen = $null
$chosenAt = [DateTime]::MinValue
$timestampSource = if (-not [string]::IsNullOrWhiteSpace($cfgBackupTimestamp)) {
    "CSV BackupTimestamp"
} else {
    "latest manifest.collectedAt"
}
foreach ($c in $candidates) {
    try {
        $m = Get-Content -Path $c.Manifest -Raw | ConvertFrom-Json
        $at = [DateTime]::MinValue
        if ($m.collectedAt) {
            $null = [DateTime]::TryParse($m.collectedAt, [ref]$at)
        }
        if ($at -eq [DateTime]::MinValue) {
            $at = (Get-Item $c.Path).LastWriteTime
        }
        if ($at -gt $chosenAt) {
            $chosen = $c
            $chosenAt = $at
        }
    }
    catch {
        # Skip unreadable manifests
    }
}

if ($null -eq $chosen) {
    Show-Error "No readable manifest.json among $($candidates.Count) candidate(s)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "No readable manifest.json")
}

$backupDir = $chosen.Path
Show-Success "Selected backup: $($chosen.PcName) / $($chosen.Timestamp)"
Show-Info "  PcName from   : $pcNameSource"
Show-Info "  Timestamp from: $timestampSource"
Show-Info "  Path: $backupDir"
if ($candidates.Count -gt 1) {
    Show-Info "  (Skipped $($candidates.Count - 1) older candidate(s))"
}
Write-Host ""


# ========================================
# Step 3b: Load & Validate Manifest
# ========================================
$manifest = $null
try {
    $manifest = Get-Content -Path $chosen.Manifest -Raw | ConvertFrom-Json
}
catch {
    Show-Error "Failed to parse manifest.json: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "manifest.json parse error")
}

if ($manifest.manifestType -ne "fabriq-printer-backup") {
    Show-Error "Unexpected manifestType: '$($manifest.manifestType)' (expected 'fabriq-printer-backup')"
    return (New-ModuleResult -Status "Error" -Message "Wrong manifestType")
}
if ([int]$manifest.schemaVersion -ne 1) {
    Show-Error "Unsupported schemaVersion: $($manifest.schemaVersion) (this module handles schemaVersion=1)"
    return (New-ModuleResult -Status "Error" -Message "Unsupported schemaVersion")
}


# ========================================
# Step 3c: Compatibility Check (Arch + OS Version)
# ========================================
$targetArch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    'arm64'
} elseif ([Environment]::Is64BitOperatingSystem) {
    'amd64'
} else {
    'x86'
}
$targetOsVersion = [System.Environment]::OSVersion.Version.ToString()

Show-Info "Compatibility check:"
Write-Host "  Backup osArch    : $($manifest.osArch)" -ForegroundColor White
Write-Host "  Target osArch    : $targetArch" -ForegroundColor White
Write-Host "  Backup osVersion : $($manifest.osVersion)" -ForegroundColor White
Write-Host "  Target osVersion : $targetOsVersion" -ForegroundColor White
Write-Host ""

if ($manifest.osArch -ne $targetArch) {
    Show-Error "Architecture mismatch: backup=$($manifest.osArch), target=$targetArch (cross-arch restore NOT supported)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Architecture mismatch")
}

$osVersionMatches = ($manifest.osVersion -eq $targetOsVersion)
if (-not $osVersionMatches) {
    if ($strictOsVersion) {
        Show-Error "osVersion mismatch and StrictOsVersion=1: refusing to proceed"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "osVersion mismatch (strict)")
    }
    else {
        Show-Warning "osVersion mismatch (StrictOsVersion=0): proceeding with caution"
        Show-Warning "  Driver / PrintConfiguration replay may produce partial fidelity"
    }
}


# ========================================
# Step 3d: Build Restore Plan
# ========================================
function Test-IsRdpRedirect {
    param($PrinterEntry)
    if ($PrinterEntry.driverName -eq 'Remote Desktop Easy Print') { return $true }
    if ($PrinterEntry.portName -match '^TS\d+$') { return $true }
    return $false
}

# Virtual / inbox-class printers that ship with Windows and almost
# never benefit from cross-PC restore. Their driver names / port names
# are stable across recent Windows builds.
$virtualDriverPatterns = @(
    'Microsoft Print To PDF',
    'Microsoft XPS Document Writer',
    'Microsoft Shared Fax Driver',
    'Microsoft OpenXPS Class Driver',
    'OneNote'
)
$virtualPortPatterns = @(
    'PORTPROMPT:',
    'XPSPort:',
    'FAX:',
    'nul:',
    'SHRFAX:'
)

function Test-IsVirtualPrinter {
    param($PrinterEntry)
    foreach ($pat in $virtualDriverPatterns) {
        if ($PrinterEntry.driverName -like "*$pat*") { return $true }
    }
    foreach ($pat in $virtualPortPatterns) {
        if ($PrinterEntry.portName -like "*$pat*") { return $true }
    }
    # Port name pattern for "Send To Microsoft OneNote" variants
    if ($PrinterEntry.portName -like 'OneNote*') { return $true }
    return $false
}

$allPrinters = @($manifest.items.printers)
$skippedRdp = @()
$skippedVirtual = @()
$plannedPrinters = @()
foreach ($p in $allPrinters) {
    if (Test-IsRdpRedirect -PrinterEntry $p) {
        $skippedRdp += $p
    }
    elseif ($skipVirtualPrinters -and (Test-IsVirtualPrinter -PrinterEntry $p)) {
        $skippedVirtual += $p
    }
    else {
        $plannedPrinters += $p
    }
}

# Ports referenced by planned printers
$referencedPortNames = @($plannedPrinters | ForEach-Object { $_.portName } | Sort-Object -Unique)

$plannedPorts = @()
foreach ($port in @($manifest.items.ports)) {
    if ($port.name -notin $referencedPortNames) { continue }
    $plannedPorts += $port
}

# Drivers referenced by planned printers
$referencedDriverNames = @($plannedPrinters | ForEach-Object { $_.driverName } | Sort-Object -Unique)

$plannedDrivers = @()
foreach ($d in @($manifest.items.drivers)) {
    if ($d.driverName -notin $referencedDriverNames) { continue }
    $plannedDrivers += $d
}

# Conflict detection
$existingPrinters = @{}
foreach ($e in @(Get-Printer -ErrorAction SilentlyContinue)) {
    $existingPrinters[$e.Name] = $e
}
$existingPorts = @{}
foreach ($e in @(Get-PrinterPort -ErrorAction SilentlyContinue)) {
    $existingPorts[$e.Name] = $e
}
$existingDriverNames = @{}
foreach ($e in @(Get-PrinterDriver -ErrorAction SilentlyContinue)) {
    $existingDriverNames[$e.Name] = $true
}

$printerConflicts = @($plannedPrinters | Where-Object { $existingPrinters.ContainsKey($_.name) })


# ========================================
# Step 3e: Display Plan
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Restore Plan" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  Source PC          : $($manifest.computerName) ($($manifest.collectedAt))" -ForegroundColor White
Write-Host "  Resolved via       : PcName=$pcNameSource, Timestamp=$timestampSource" -ForegroundColor DarkGray
Write-Host "  Target PC          : $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  OS Version Match   : $(if ($osVersionMatches) { 'YES' } else { 'NO (warning, proceeding)' })" -ForegroundColor White
Write-Host ""
Write-Host "  Printers to restore: $($plannedPrinters.Count) (of $($allPrinters.Count))" -ForegroundColor White
Write-Host "  Skipped (RDP-redir): $($skippedRdp.Count)" -ForegroundColor DarkGray
Write-Host "  Skipped (virtual)  : $($skippedVirtual.Count) (SkipVirtualPrinters=$(if ($skipVirtualPrinters) { '1' } else { '0' }))" -ForegroundColor DarkGray
Write-Host "  Ports to ensure    : $($plannedPorts.Count)" -ForegroundColor White
Write-Host "  Drivers to ensure  : $($plannedDrivers.Count)" -ForegroundColor White
Write-Host "  Existing conflicts : $($printerConflicts.Count) (OnConflict=$onConflict)" -ForegroundColor White
Write-Host "  ReuseInboxDrivers  : $(if ($reuseInboxDrivers) { 'YES' } else { 'NO (force payload install)' })" -ForegroundColor White
Write-Host "  Set default printer: $(if ($restoreDefaultPrinter) { 'YES' } else { 'NO' })" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($plannedPrinters.Count -eq 0) {
    Show-Info "No printers to restore after filtering"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Nothing to restore after RDP filter")
}


# ========================================
# Step 4: Confirm
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with printer restore?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5a: Restore Drivers
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[Phase 1/4] Restoring drivers" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$driverNameToStorePath = @{}  # driverName -> InfPath in DriverStore (for Add-PrinterDriver)
$driverSuccess = 0
$driverSkip    = 0
$driverFail    = 0
$warnings      = @()

# Group planned drivers by their backup payload folder (one pnputil call per oemNN.inf)
$payloadGroups = @{}
foreach ($d in $plannedDrivers) {
    $key = if ($d.backupFolder) { $d.backupFolder } else { "__no_payload__" }
    if (-not $payloadGroups.ContainsKey($key)) { $payloadGroups[$key] = @() }
    $payloadGroups[$key] += $d
}

# Per-payload-folder install + per-driver-name Add-PrinterDriver
foreach ($key in @($payloadGroups.Keys)) {
    $drivers = $payloadGroups[$key]
    $firstDriver = $drivers[0]

    # Decide whether to re-install payload or rely on existing/inbox driver
    $useExisting = $false
    if ($key -eq "__no_payload__") {
        # No payload in backup -> must rely on existing/inbox
        $useExisting = $true
    }
    elseif ($reuseInboxDrivers -and ($firstDriver.isInboxDriver -or $firstDriver.manufacturer -eq 'Microsoft')) {
        # ReuseInboxDrivers policy: skip payload for Microsoft-supplied drivers
        $useExisting = $true
    }

    $storeInfPath = $null

    if (-not $useExisting) {
        $payloadDir = Join-Path $backupDir $firstDriver.backupFolder
        if (-not (Test-Path $payloadDir)) {
            Show-Warning "Payload missing for '$($firstDriver.backupFolder)' — falling back to existing driver"
            $warnings += "Payload missing: $($firstDriver.backupFolder)"
            $useExisting = $true
        }
        else {
            # Find the actual .inf file inside the payload folder (one expected)
            $infFile = @(Get-ChildItem -Path $payloadDir -Filter *.inf -File -ErrorAction SilentlyContinue) | Select-Object -First 1
            if ($null -eq $infFile) {
                Show-Warning "No .inf file inside payload: $payloadDir — falling back to existing"
                $warnings += "No .inf in payload: $($firstDriver.backupFolder)"
                $useExisting = $true
            }
            else {
                Show-Info "  pnputil /add-driver $($infFile.Name)"
                $null = & pnputil /add-driver $infFile.FullName /install 2>&1
                $exitCode = $LASTEXITCODE
                # pnputil returns 0 success or 259 (no more data) when driver already exists
                if ($exitCode -ne 0 -and $exitCode -ne 259) {
                    Show-Warning "    pnputil exit code: $exitCode (continuing; may already be present)"
                    $warnings += "pnputil exit $exitCode for $($infFile.Name)"
                }

                # Resolve resulting DriverStore folder path
                $infBase = [System.IO.Path]::GetFileNameWithoutExtension($infFile.Name).ToLower()
                $repo = "C:\Windows\System32\DriverStore\FileRepository"
                $storeDir = @(Get-ChildItem -Path $repo -Directory -Filter "${infBase}.inf_${targetArch}_*" -ErrorAction SilentlyContinue |
                              Sort-Object LastWriteTime -Descending) | Select-Object -First 1
                if ($null -eq $storeDir) {
                    Show-Warning "    DriverStore folder not found for $infBase — falling back to existing"
                    $warnings += "DriverStore folder missing for $infBase"
                    $useExisting = $true
                }
                else {
                    $storeInfPath = Join-Path $storeDir.FullName $infFile.Name
                    if (-not (Test-Path $storeInfPath)) {
                        Show-Warning "    Store INF not found: $storeInfPath — falling back to existing"
                        $warnings += "Store INF missing: $storeInfPath"
                        $useExisting = $true
                    }
                }
            }
        }
    }

    # Register each driver name from this group
    foreach ($d in $drivers) {
        if ($existingDriverNames.ContainsKey($d.driverName)) {
            Show-Skip "Driver already registered: $($d.driverName)"
            $driverNameToStorePath[$d.driverName] = $null  # already present
            $driverSkip++
            continue
        }

        if ($useExisting) {
            # No payload path available — attempt Add-PrinterDriver -Name only
            # (works only if Windows already knows this driver from inbox)
            try {
                Add-PrinterDriver -Name $d.driverName -ErrorAction Stop
                Show-Success "Driver registered (inbox/existing): $($d.driverName)"
                $driverSuccess++
            }
            catch {
                Show-Error "Driver not available without payload: $($d.driverName) — $($_.Exception.Message)"
                $warnings += "Inbox-only restore failed for driver: $($d.driverName)"
                $driverFail++
            }
            continue
        }

        try {
            Add-PrinterDriver -Name $d.driverName -InfPath $storeInfPath -ErrorAction Stop
            Show-Success "Driver registered: $($d.driverName)"
            $driverNameToStorePath[$d.driverName] = $storeInfPath
            $driverSuccess++
        }
        catch {
            Show-Error "Add-PrinterDriver failed: $($d.driverName) — $($_.Exception.Message)"
            $warnings += "Add-PrinterDriver failed: $($d.driverName)"
            $driverFail++
        }
    }
}

Write-Host ""


# ========================================
# Step 5b: Restore Ports
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[Phase 2/4] Restoring ports" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$portSuccess = 0
$portSkip    = 0
$portFail    = 0

foreach ($port in $plannedPorts) {
    if ($existingPorts.ContainsKey($port.name)) {
        Show-Skip "Port already exists: $($port.name)"
        $portSkip++
        continue
    }

    switch ($port.portType) {
        'TCPIP' {
            if ([string]::IsNullOrWhiteSpace($port.printerHostAddress)) {
                Show-Error "TCPIP port missing printerHostAddress: $($port.name)"
                $warnings += "Port skipped (no host address): $($port.name)"
                $portFail++
                break
            }
            try {
                $params = @{
                    Name               = $port.name
                    PrinterHostAddress = $port.printerHostAddress
                    ErrorAction        = 'Stop'
                }
                if ($port.portNumber) { $params['PortNumber'] = [int]$port.portNumber }
                Add-PrinterPort @params
                Show-Success "TCPIP port created: $($port.name) -> $($port.printerHostAddress):$($port.portNumber)"
                $portSuccess++
            }
            catch {
                Show-Error "Failed to create TCPIP port $($port.name): $($_.Exception.Message)"
                $warnings += "TCPIP port failed: $($port.name)"
                $portFail++
            }
        }
        'LPR' {
            if ([string]::IsNullOrWhiteSpace($port.lprHostName)) {
                Show-Error "LPR port missing lprHostName: $($port.name)"
                $warnings += "Port skipped (no LPR host): $($port.name)"
                $portFail++
                break
            }
            try {
                Add-PrinterPort -Name $port.name -LprHostName $port.lprHostName -LprQueueName $port.lprQueueName -ErrorAction Stop
                Show-Success "LPR port created: $($port.name)"
                $portSuccess++
            }
            catch {
                Show-Error "Failed to create LPR port $($port.name): $($_.Exception.Message)"
                $warnings += "LPR port failed: $($port.name)"
                $portFail++
            }
        }
        'Local' {
            # Standard local ports (LPT1:, COM1:, etc.) are Windows built-ins; skip
            Show-Skip "Standard local port (system-provided): $($port.name)"
            $portSkip++
        }
        'WSD' {
            Show-Warning "WSD port not restorable (dynamic discovery): $($port.name)"
            $warnings += "WSD port skipped: $($port.name)"
            $portSkip++
        }
        'Bonjour' {
            Show-Warning "Bonjour port not auto-restorable: $($port.name)"
            $warnings += "Bonjour port skipped: $($port.name)"
            $portSkip++
        }
        default {
            Show-Warning "Unsupported port type '$($port.portType)': $($port.name) (skipped)"
            $warnings += "Port type '$($port.portType)' skipped: $($port.name)"
            $portSkip++
        }
    }
}

Write-Host ""


# ========================================
# Step 5c: Restore Printers
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[Phase 3/4] Restoring printers" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$printerSuccess = 0
$printerSkip    = 0
$printerFail    = 0
$restoredPrinterNames = @()

foreach ($p in $plannedPrinters) {
    if ($existingPrinters.ContainsKey($p.name)) {
        if ($onConflict -eq "skip") {
            Show-Skip "Printer already exists (OnConflict=skip): $($p.name)"
            $printerSkip++
            continue
        }
        # OnConflict=replace
        try {
            Remove-Printer -Name $p.name -ErrorAction Stop
            Show-Info "Removed existing printer (OnConflict=replace): $($p.name)"
        }
        catch {
            Show-Error "Failed to remove existing printer $($p.name): $($_.Exception.Message)"
            $warnings += "Replace failed for $($p.name)"
            $printerFail++
            continue
        }
    }

    try {
        Add-Printer -Name $p.name -DriverName $p.driverName -PortName $p.portName -ErrorAction Stop
        Show-Success "Printer created: $($p.name) (driver: $($p.driverName), port: $($p.portName))"
        $printerSuccess++
        $restoredPrinterNames += $p.name

        # Apply printer-level attributes (best-effort)
        try {
            if ($p.shared -and -not [string]::IsNullOrWhiteSpace($p.shareName)) {
                Set-Printer -Name $p.name -Shared $true -ShareName $p.shareName -ErrorAction Stop
            }
            if (-not [string]::IsNullOrWhiteSpace($p.comment)) {
                Set-Printer -Name $p.name -Comment $p.comment -ErrorAction SilentlyContinue
            }
            if (-not [string]::IsNullOrWhiteSpace($p.location)) {
                Set-Printer -Name $p.name -Location $p.location -ErrorAction SilentlyContinue
            }
        }
        catch {
            $warnings += "Printer attribute set failed for '$($p.name)': $($_.Exception.Message)"
        }
    }
    catch {
        Show-Error "Add-Printer failed: $($p.name) — $($_.Exception.Message)"
        $warnings += "Add-Printer failed: $($p.name)"
        $printerFail++
    }
}

Write-Host ""


# ========================================
# Step 5d: Restore Print Settings (PrintTicketXML round-trip)
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "[Phase 4/4] Restoring print settings" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Resolve target HKCU hive once. Under UAC elevation from a different
# account, or SYSTEM context, bare `HKCU:` points to the WRONG user.
# Resolve-HkcuRoot redirects to HKU\<LoggedOnUserSid> when needed
# (same pattern as reg_hkcu_config / desktop_icon_restore).
$hkcuInfo = Resolve-HkcuRoot
if ($hkcuInfo.Redirected) {
    Show-Info "Per-user DEVMODE target: $($hkcuInfo.Label) [SID=$($hkcuInfo.SID)]"
}
$devModeKey = $hkcuInfo.PsDrivePath + '\Printers\DevModePerUser'

$settingsSuccess = 0
$settingsSkip    = 0
$settingsFail    = 0
$hwConfigRestoredCount = 0

foreach ($p in $plannedPrinters) {
    if ($p.name -notin $restoredPrinterNames -and -not $existingPrinters.ContainsKey($p.name)) {
        $settingsSkip++
        continue
    }

    # PrintConfiguration (PrintTicketXML round-trip + explicit public DEVMODE
    # field overrides). The round-trip alone is insufficient because some
    # drivers ignore DM_COLOR/DM_DEFAULTSOURCE when parsing PrintTicketXML
    # (they prefer their private DEVMODE namespace). Explicit -Color /
    # -DuplexingMode / -PaperSize / -PaperSource / -PrintQuality / -Collate
    # calls force the public DEVMODE fields directly via the spooler API.
    if (-not [string]::IsNullOrWhiteSpace($p.printSettingsFile)) {
        $xmlPath = Join-Path $backupDir $p.printSettingsFile
        if (Test-Path $xmlPath) {
            try {
                $cfgObj = Import-Clixml -Path $xmlPath -ErrorAction Stop
                $pt = $cfgObj.PrintTicketXML
                if (-not [string]::IsNullOrWhiteSpace($pt)) {
                    Set-PrintConfiguration -PrinterName $p.name -PrintTicketXml $pt -ErrorAction Stop
                    Show-Success "Print configuration restored: $($p.name)"
                    $settingsSuccess++
                }
                else {
                    $warnings += "PrintTicketXML empty for '$($p.name)' - skipped"
                    $settingsSkip++
                }

                # Explicit field overrides. Each in its own try/catch since
                # any individual one may be unsupported by the driver.
                $explicitFields = @(
                    @{ Name = 'Color';         Parser = { [System.Convert]::ToBoolean($cfgObj.Color) } }
                    @{ Name = 'Collate';       Parser = { [System.Convert]::ToBoolean($cfgObj.Collate) } }
                    @{ Name = 'DuplexingMode'; Parser = { "$($cfgObj.DuplexingMode)" } }
                    @{ Name = 'PaperSize';     Parser = { "$($cfgObj.PaperSize)" } }
                    @{ Name = 'PaperSource';   Parser = { "$($cfgObj.PaperSource)" } }
                    @{ Name = 'PrintQuality';  Parser = { "$($cfgObj.PrintQuality)" } }
                )
                $appliedFields = @()
                foreach ($field in $explicitFields) {
                    $raw = $cfgObj.($field.Name)
                    if ($null -eq $raw -or "$raw" -eq '') { continue }
                    try {
                        $val = & $field.Parser
                        $params = @{ PrinterName = $p.name; $field.Name = $val; ErrorAction = 'Stop' }
                        Set-PrintConfiguration @params
                        $appliedFields += "$($field.Name)=$val"
                    }
                    catch {
                        # Field-level failures are normal (driver-specific support);
                        # tally as warning but don't break the printer.
                        $warnings += "Set-PrintConfiguration -$($field.Name) failed for '$($p.name)': $($_.Exception.Message)"
                    }
                }
                if ($appliedFields.Count -gt 0) {
                    Show-Info "  explicit fields: $($appliedFields -join ', ')"
                }
            }
            catch {
                $warnings += "Set-PrintConfiguration failed for '$($p.name)': $($_.Exception.Message)"
                Show-Warning "Print configuration failed (continuing): $($p.name)"
                $settingsFail++
            }
        }
        else {
            $settingsSkip++
        }
    }

    # Printer properties (best-effort, many are read-only)
    if (-not [string]::IsNullOrWhiteSpace($p.propertiesFile)) {
        $propPath = Join-Path $backupDir $p.propertiesFile
        if (Test-Path $propPath) {
            try {
                $props = Get-Content -Path $propPath -Raw | ConvertFrom-Json
                foreach ($prop in @($props)) {
                    if ([string]::IsNullOrWhiteSpace($prop.PropertyName)) { continue }
                    try {
                        Set-PrinterProperty -PrinterName $p.name -PropertyName $prop.PropertyName -Value $prop.Value -ErrorAction Stop
                    }
                    catch {
                        # Most properties are read-only; silent skip is normal
                    }
                }
            }
            catch {
                $warnings += "Property file read failed for '$($p.name)': $($_.Exception.Message)"
            }
        }
    }

    # DevModePerUser and HwConfig writes are deferred to pass 2 (after
    # Spooler restart) so that our values land as the last writes the
    # driver sees, avoiding the driver's post-restart re-init from
    # overwriting them.
}

# Pass 1 (Set-PrintConfiguration + explicit + Set-PrinterProperty) is now
# complete. Decide whether to restart Spooler before pass 2.
$anyHwConfigPlanned = $restoreHardwareConfig -and
    @($plannedPrinters | Where-Object { -not [string]::IsNullOrWhiteSpace($_.hwConfigFile) }).Count -gt 0
$anyDevModePlanned = @($plannedPrinters | Where-Object { -not [string]::IsNullOrWhiteSpace($_.devModeFile) }).Count -gt 0

if ($anyHwConfigPlanned -or $anyDevModePlanned) {
    # Pre-write Spooler restart. Lets the driver settle post-Add-Printer
    # and post-Set-PrintConfiguration. After this, our pass-2 writes are
    # the freshest state. Without this, the driver re-init that happens
    # ON Spooler restart can overwrite our binary blobs.
    Show-Info "Restarting Print Spooler to stabilize before HW config / DEVMODE writes..."
    try {
        Restart-Service -Name Spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Show-Success "Spooler restarted"
    }
    catch {
        Show-Warning "Spooler restart failed: $($_.Exception.Message)"
        $warnings += "Spooler restart failed: $($_.Exception.Message)"
    }
}

# Pass 2: Write HKLM PrinterDriverData (installable options / vendor
# binary blobs) and HKCU DevModePerUser (per-user print dialog default)
# AFTER Spooler restart so the driver does not auto-reset our writes
# during its post-restart re-initialization.
foreach ($p in $plannedPrinters) {
    if ($p.name -notin $restoredPrinterNames -and -not $existingPrinters.ContainsKey($p.name)) {
        continue
    }

    # Hardware / installable-options config (HKLM-side, system-wide).
    # Vendor-specific binary blobs that determine which paper trays /
    # finishers / duplex units the print UI shows. Without this, the
    # target printer falls back to the driver's minimum hardware config
    # (e.g. only Cassette 1 visible).
    # Risk: vendor blob format is driver-version-specific. Cross-version
    # restore may corrupt the blob and break printer state. Opt-in via
    # RestoreHardwareConfig=1.
    if ($restoreHardwareConfig -and -not [string]::IsNullOrWhiteSpace($p.hwConfigFile)) {
        $hwConfigPath = Join-Path $backupDir $p.hwConfigFile
        if (Test-Path $hwConfigPath) {
            $hwRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($p.name)\PrinterDriverData"
            try {
                $hwDump = Get-Content -Path $hwConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json
                if (-not (Test-Path $hwRegPath)) {
                    $null = New-Item -Path $hwRegPath -Force -ErrorAction Stop
                }
                $restoredValues = 0
                foreach ($prop in $hwDump.PSObject.Properties) {
                    $vname = $prop.Name
                    $info = $prop.Value
                    if ($null -eq $info -or [string]::IsNullOrWhiteSpace($info.Type)) { continue }
                    try {
                        $decoded = switch ($info.Type) {
                            'Binary'       { if ($null -ne $info.Data) { [Convert]::FromBase64String([string]$info.Data) } else { [byte[]]@() } }
                            'String'       { [string]$info.Data }
                            'ExpandString' { [string]$info.Data }
                            'MultiString'  { @($info.Data) }
                            'DWord'        { [int]$info.Data }
                            'QWord'        { [long]$info.Data }
                            default        { [string]$info.Data }
                        }
                        $null = New-ItemProperty -Path $hwRegPath -Name $vname -Value $decoded `
                            -PropertyType $info.Type -Force -ErrorAction Stop
                        $restoredValues++
                    }
                    catch {
                        $warnings += "HwConfig value '$vname' on '$($p.name)' failed: $($_.Exception.Message)"
                    }
                }
                if ($restoredValues -gt 0) {
                    Show-Success "Hardware config restored: $($p.name) ($restoredValues value(s))"
                    $hwConfigRestoredCount++
                }
            }
            catch {
                $warnings += "Hardware config restore failed for '$($p.name)': $($_.Exception.Message)"
                Show-Warning "Hardware config restore failed (continuing): $($p.name)"
            }
        }
    }

    # Per-user DEVMODE blob (HKCU). Last write so driver re-init can not
    # overwrite. Resolve-HkcuRoot resolution already done before pass 1.
    if (-not [string]::IsNullOrWhiteSpace($p.devModeFile)) {
        $devModePath = Join-Path $backupDir $p.devModeFile
        if (Test-Path $devModePath) {
            try {
                $b64 = (Get-Content -Path $devModePath -Raw -ErrorAction Stop).Trim()
                if (-not [string]::IsNullOrWhiteSpace($b64)) {
                    $blob = [Convert]::FromBase64String($b64)
                    if (-not (Test-Path $devModeKey)) {
                        $null = New-Item -Path $devModeKey -Force -ErrorAction Stop
                    }
                    $null = New-ItemProperty -Path $devModeKey -Name $p.name -Value $blob `
                        -PropertyType Binary -Force -ErrorAction Stop
                    Show-Success "Per-user DEVMODE restored: $($p.name)"
                }
            }
            catch {
                $warnings += "Per-user DEVMODE restore failed for '$($p.name)': $($_.Exception.Message)"
                Show-Warning "Per-user DEVMODE restore failed (continuing): $($p.name)"
            }
        }
    }
}

Write-Host ""


# ========================================
# Step 5e: Set Default Printer
# ========================================
if ($restoreDefaultPrinter -and -not [string]::IsNullOrWhiteSpace($manifest.defaultPrinter)) {
    $defName = $manifest.defaultPrinter

    # Only attempt to set default if this printer was in the restored (planned) set.
    # Printers filtered out as RDP redirects are skipped here automatically.
    $defWasPlanned = @($plannedPrinters | Where-Object { $_.name -eq $defName }).Count -gt 0

    if (-not $defWasPlanned) {
        Show-Info "Default printer '$defName' was not in the restored set (likely RDP-redirected); not setting"
    }
    else {
        $defExists = $null -ne (Get-Printer -Name $defName -ErrorAction SilentlyContinue)
        if ($defExists) {
            try {
                $shell = New-Object -ComObject WScript.Network
                $shell.SetDefaultPrinter($defName)
                Show-Success "Default printer set: $defName"
            }
            catch {
                Show-Warning "Failed to set default printer '$defName': $($_.Exception.Message)"
                $warnings += "SetDefaultPrinter failed: $defName"
            }
        }
        else {
            Show-Warning "Default printer '$defName' not present after restore (skipping)"
            $warnings += "Default printer absent post-restore: $defName"
        }
    }
}

Write-Host ""


# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Post-Apply Verification" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$verifyPass = 0
$verifyFail = 0

foreach ($p in $plannedPrinters) {
    $actual = Get-Printer -Name $p.name -ErrorAction SilentlyContinue
    if ($null -eq $actual) {
        if ($onConflict -eq "skip" -and $existingPrinters.ContainsKey($p.name)) {
            # Was a skipped pre-existing one that mysteriously disappeared
            Write-Host "  [VERIFY FAILED] $($p.name) - expected to exist (was skipped)" -ForegroundColor Red
            $verifyFail++
        }
        elseif ($p.name -notin $restoredPrinterNames) {
            # We never tried to create it (creation failed in step 5c)
            Write-Host "  [VERIFY SKIPPED] $($p.name) - not created" -ForegroundColor DarkGray
        }
        else {
            Write-Host "  [VERIFY FAILED] $($p.name) - not found post-restore" -ForegroundColor Red
            $verifyFail++
        }
        continue
    }

    $driverOk = $actual.DriverName -eq $p.driverName
    $portOk   = $actual.PortName   -eq $p.portName

    if ($driverOk -and $portOk) {
        Write-Host "  [VERIFIED] $($p.name)" -ForegroundColor Green
        $verifyPass++
    }
    else {
        $reasons = @()
        if (-not $driverOk) { $reasons += "driver mismatch ($($actual.DriverName))" }
        if (-not $portOk)   { $reasons += "port mismatch ($($actual.PortName))" }
        Write-Host "  [VERIFY FAILED] $($p.name) - $($reasons -join ', ')" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0 -and $printerFail -eq 0)


# ========================================
# Step 6: Result Summary
# ========================================
Show-Separator
Write-Host "Printer Restore Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Source       : $($manifest.computerName) / $($chosen.Timestamp)" -ForegroundColor White
Write-Host "  Drivers      : $driverSuccess success / $driverSkip skip / $driverFail fail" -ForegroundColor White
Write-Host "  Ports        : $portSuccess success / $portSkip skip / $portFail fail" -ForegroundColor White
Write-Host "  Printers     : $printerSuccess success / $printerSkip skip / $printerFail fail" -ForegroundColor White
Write-Host "  Print config : $settingsSuccess success / $settingsSkip skip / $settingsFail fail" -ForegroundColor White
Write-Host "  HW config    : $hwConfigRestoredCount printer(s) restored (RestoreHardwareConfig=$(if ($restoreHardwareConfig) { '1' } else { '0' }))" -ForegroundColor White
Write-Host "  Skipped RDP  : $($skippedRdp.Count)" -ForegroundColor DarkGray
Write-Host "  Skipped virt : $($skippedVirtual.Count)" -ForegroundColor DarkGray
if ($warnings.Count -gt 0) {
    Write-Host "  Warnings     : $($warnings.Count)" -ForegroundColor Yellow
}
if ($verified) {
    Write-Host "  Verified     : PASS" -ForegroundColor Green
} else {
    Write-Host "  Verified     : FAIL" -ForegroundColor Red
}
Show-Separator
Write-Host ""

$totalFail = $printerFail + $driverFail + $portFail
$totalSuccess = $printerSuccess

return (New-BatchResult `
    -Success $totalSuccess `
    -Skip ($printerSkip + $skippedRdp.Count + $skippedVirtual.Count) `
    -Fail $totalFail `
    -Title "Printer Restore Results" `
    -MessageSuffix " (drivers: $driverSuccess, ports: $portSuccess, settings: $settingsSuccess)" `
    -Verified $verified)
