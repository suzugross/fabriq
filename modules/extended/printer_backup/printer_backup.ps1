# ========================================
# Printer Backup
# ========================================
# [PURPOSE]
# Capture this PC's printer drivers, ports, printers, default printer,
# and per-printer print settings into a portable, versioned backup
# folder that a companion restore module can later replay.
#
# [NOTES]
# - Requires administrator privileges (pnputil /export-driver,
#   Get-PrinterDriver, Get-WindowsDriver -Online).
# - Output: backup/<ComputerName>/<yyyy_MM_dd_HHmmss>/
# - Inbox printer drivers (Microsoft Print to PDF / OneNote / XPS / Fax)
#   are recorded in the manifest but their payload is NOT exported
#   (Windows ships them; restore re-uses the in-box copy).
# - WSD ports are recorded but flagged as restore-unsafe (depend on
#   dynamic discovery at restore time).
# - Driver payload bytes are large (50-200 MB per package). Set
#   IncludeDriverBinaries=0 in the config CSV to skip payload export
#   when only metadata is needed.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Printer Backup" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Step 1: Load Config CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "printer_backup_config.csv"

$cfgItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "IncludeDriverBinaries", "IncludePrintSettings")

if ($null -eq $cfgItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load printer_backup_config.csv")
}
if (@($cfgItems).Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "Backup disabled in config CSV")
}

$cfg = @($cfgItems)[0]
$includeDriverBinaries = ($cfg.IncludeDriverBinaries -eq "1")
$includePrintSettings  = ($cfg.IncludePrintSettings  -eq "1")


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
# Step 3: Scan & Dry-Run Summary
# ========================================
Show-Info "Scanning printer environment..."

$printers = @()
try {
    $printers = @(Get-Printer -ErrorAction Stop)
}
catch {
    Show-Error "Get-Printer failed: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Get-Printer failed")
}

if ($printers.Count -eq 0) {
    Show-Info "No printers found on this PC"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "No printers to back up")
}

$ports = @()
try {
    $ports = @(Get-PrinterPort -ErrorAction Stop)
}
catch {
    Show-Warning "Get-PrinterPort failed: $_"
}

$printerDrivers = @()
try {
    $printerDrivers = @(Get-PrinterDriver -ErrorAction Stop)
}
catch {
    Show-Warning "Get-PrinterDriver failed: $_"
}

# Third-party INF packages (oemNN.inf). Get-WindowsDriver -Online only
# enumerates 3rd-party drivers; inbox drivers do not appear here, which
# is exactly the discriminator we use for IsInboxDriver below.
$thirdPartyInfs = @()
try {
    $thirdPartyInfs = @(
        Get-WindowsDriver -Online -ErrorAction Stop |
        Where-Object { $_.ClassName -eq 'Printer' }
    )
}
catch {
    Show-Warning "Get-WindowsDriver failed: $_"
}

# basename (lower-case, no extension) -> WindowsDriver object
$basenameToOem = @{}
foreach ($wd in $thirdPartyInfs) {
    $orig = $wd.OriginalFileName
    if (-not [string]::IsNullOrWhiteSpace($orig)) {
        $bn = [System.IO.Path]::GetFileNameWithoutExtension($orig).ToLower()
        if (-not $basenameToOem.ContainsKey($bn)) {
            $basenameToOem[$bn] = $wd
        }
    }
}

# Per-registered-driver classification + payload export plan
$driverInfoList = @()
$infPackagesToExport = @{}
foreach ($pd in $printerDrivers) {
    $bn = if ($pd.InfPath) { [System.IO.Path]::GetFileNameWithoutExtension($pd.InfPath).ToLower() } else { $null }
    $isThird = ($null -ne $bn) -and $basenameToOem.ContainsKey($bn)
    $oemInf = $null
    if ($isThird) {
        $wd = $basenameToOem[$bn]
        $oemInf = $wd.Driver
        if ($includeDriverBinaries -and -not $infPackagesToExport.ContainsKey($oemInf)) {
            $infPackagesToExport[$oemInf] = $wd
        }
    }
    $driverInfoList += [PSCustomObject]@{
        DriverName    = $pd.Name
        Manufacturer  = $pd.Manufacturer
        DriverVersion = $pd.DriverVersion
        InfPath       = $pd.InfPath
        InfBaseName   = $bn
        OemInf        = $oemInf
        IsInboxDriver = -not $isThird
    }
}

# Resolve PC name for backup folder identity. Prefer hostlist OldPCname
# (kitting workflow's view of the customer's PC) and fall back to the
# current Windows COMPUTERNAME when no hostlist is in effect.
$pcName = if (-not [string]::IsNullOrWhiteSpace($env:SELECTED_OLD_PCNAME)) {
    $env:SELECTED_OLD_PCNAME.Trim()
} else {
    $env:COMPUTERNAME
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Backup Plan" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  PC name (folder):   $pcName" -ForegroundColor White
if ($pcName -ne $env:COMPUTERNAME) {
    Write-Host "  Computer name:      $env:COMPUTERNAME (current Windows hostname)" -ForegroundColor DarkGray
}
Write-Host "  Printers:           $($printers.Count)" -ForegroundColor White
Write-Host "  Ports:              $($ports.Count)" -ForegroundColor White
Write-Host "  Registered drivers: $($printerDrivers.Count)" -ForegroundColor White
$payloadNote = if (-not $includeDriverBinaries) { ' (NOT exported, IncludeDriverBinaries=0)' } else { '' }
Write-Host "  3rd-party packages: $($infPackagesToExport.Count)$payloadNote" -ForegroundColor White
$psNote = if ($includePrintSettings) { 'Yes' } else { 'No' }
Write-Host "  Print settings:     $psNote" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: Confirm
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with printer backup?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: Execute Backup
# ========================================
# $pcName was resolved in Step 3 (OldPCname preferred, fallback to COMPUTERNAME)
$timestamp = Get-Date -Format "yyyy_MM_dd_HHmmss"
$backupDir = Join-Path (Join-Path (Join-Path $PSScriptRoot "backup") $pcName) $timestamp

try {
    $null = New-Item -ItemType Directory -Path $backupDir -Force -ErrorAction Stop
}
catch {
    Show-Error "Failed to create backup directory: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to create backup dir")
}

Show-Info "Backup target: $backupDir"
Write-Host ""

$warnings    = @()
$failCount   = 0


function Write-JsonArray {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        $Data
    )
    # ConvertTo-Json with -InputObject avoids pipeline unrolling so a
    # single-element array still serializes as a JSON array, not an object.
    $json = ConvertTo-Json -InputObject @($Data) -Depth 6
    $json | Out-File -FilePath $Path -Encoding UTF8 -Force
}


# ---- 5.1: printers.json
Show-Info "Writing printers.json..."
try {
    $printerRows = $printers | Select-Object Name, DriverName, PortName, Shared, ShareName, Comment, Location, Published, KeepPrintedJobs
    Write-JsonArray -Path (Join-Path $backupDir "printers.json") -Data $printerRows
    Show-Success "printers.json written"
}
catch {
    Show-Error "Failed to write printers.json: $_"
    $failCount++
}


# ---- 5.2: ports.json (classified by PortType)
Show-Info "Writing ports.json..."

function Get-PortType {
    param($Port)
    $monitor = $Port.PortMonitor
    if ([string]::IsNullOrEmpty($monitor)) { return 'Other' }
    $m = $monitor.ToLower()
    if ($m -like 'tcpmon*')   { return 'TCPIP' }
    if ($m -like 'lprmon*' -or $m -like 'lpr*') { return 'LPR' }
    if ($m -like 'wsd*')      { return 'WSD' }
    if ($m -like 'localmon*' -or $m -like 'local*') { return 'Local' }
    if ($m -like '*bonjour*' -or $m -like '*mdns*') { return 'Bonjour' }
    return 'Other'
}

try {
    $portRows = foreach ($p in $ports) {
        [PSCustomObject]@{
            Name               = $p.Name
            PortType           = Get-PortType -Port $p
            Description        = $p.Description
            PortMonitor        = $p.PortMonitor
            PrinterHostAddress = $p.PrinterHostAddress
            PortNumber         = $p.PortNumber
            LprHostName        = $p.LprHostName
            LprQueueName       = $p.LprQueueName
            SnmpEnabled        = [bool]$p.SnmpEnabled
            SnmpCommunity      = $p.SnmpCommunity
        }
    }
    Write-JsonArray -Path (Join-Path $backupDir "ports.json") -Data $portRows
    Show-Success "ports.json written"

    foreach ($w in @($portRows | Where-Object { $_.PortType -eq 'WSD' })) {
        $warnings += "WSD port '$($w.Name)': dynamic discovery dependent, restore not guaranteed"
    }
}
catch {
    Show-Error "Failed to write ports.json: $_"
    $failCount++
}


# ---- 5.3: drivers_registered.json
Show-Info "Writing drivers_registered.json..."
try {
    Write-JsonArray -Path (Join-Path $backupDir "drivers_registered.json") -Data $driverInfoList
    Show-Success "drivers_registered.json written"
}
catch {
    Show-Error "Failed to write drivers_registered.json: $_"
    $failCount++
}


# ---- 5.4: drivers_inf_inventory.json (third-party only)
Show-Info "Writing drivers_inf_inventory.json..."
try {
    $invRows = foreach ($wd in $thirdPartyInfs) {
        [PSCustomObject]@{
            Driver           = $wd.Driver
            OriginalFileName = $wd.OriginalFileName
            ClassName        = $wd.ClassName
            ProviderName     = $wd.ProviderName
            Date             = if ($wd.Date) { $wd.Date.ToString("o") } else { $null }
            Version          = $wd.Version
            Inbox            = [bool]$wd.Inbox
        }
    }
    Write-JsonArray -Path (Join-Path $backupDir "drivers_inf_inventory.json") -Data $invRows
    Show-Success "drivers_inf_inventory.json written"
}
catch {
    Show-Error "Failed to write drivers_inf_inventory.json: $_"
    $failCount++
}


# ---- 5.5: per-printer print settings (optional)
function Get-SafeFileName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "unnamed" }
    return ($Name -replace '[^\w\-]', '_')
}

$printSettingsRefs = @{}

if ($includePrintSettings) {
    Show-Info "Capturing per-printer print settings..."
    $settingsDir = Join-Path $backupDir "printsettings"
    $null = New-Item -ItemType Directory -Path $settingsDir -Force -ErrorAction SilentlyContinue

    # Resolve which user's HKCU we are actually targeting. Under UAC
    # elevation from a different account, or under SYSTEM context (e.g.
    # RunOnce / Scheduled Task), bare `HKCU:` points to the WRONG user.
    # Resolve-HkcuRoot redirects to HKU\<LoggedOnUserSid> when needed
    # (same pattern as reg_hkcu_config / desktop_icon_backup).
    $hkcuInfo = Resolve-HkcuRoot
    if ($hkcuInfo.Redirected) {
        Show-Info "Per-user DEVMODE source: $($hkcuInfo.Label) [SID=$($hkcuInfo.SID)]"
    }
    $devModeKey = $hkcuInfo.PsDrivePath + '\Printers\DevModePerUser'

    $usedNames = @{}

    foreach ($p in $printers) {
        $safe = Get-SafeFileName -Name $p.Name
        if ($usedNames.ContainsKey($safe)) {
            $usedNames[$safe]++
            $safe = "${safe}_$($usedNames[$safe])"
        } else {
            $usedNames[$safe] = 1
        }

        $xmlFile      = "printsettings/${safe}.xml"
        $propFile     = "printsettings/${safe}.properties.json"
        $devModeFile  = "printsettings/${safe}.devmode.b64"
        $hwConfigFile = "printsettings/${safe}.hwconfig.json"
        $xmlPath      = Join-Path $backupDir $xmlFile
        $propPath     = Join-Path $backupDir $propFile
        $devModePath  = Join-Path $backupDir $devModeFile
        $hwConfigPath = Join-Path $backupDir $hwConfigFile

        $xmlOk      = $false
        $propOk     = $false
        $devModeOk  = $false
        $hwConfigOk = $false

        try {
            $config = Get-PrintConfiguration -PrinterName $p.Name -ErrorAction Stop
            $config | Export-Clixml -Path $xmlPath -Force -ErrorAction Stop
            $xmlOk = $true
        }
        catch {
            $warnings += "Print configuration capture failed for '$($p.Name)': $($_.Exception.Message)"
        }

        try {
            $props = Get-PrinterProperty -PrinterName $p.Name -ErrorAction Stop
            $propRows = $props | Select-Object PrinterName, PropertyName, Type, Value
            Write-JsonArray -Path $propPath -Data $propRows
            $propOk = $true
        }
        catch {
            $warnings += "Printer property capture failed for '$($p.Name)': $($_.Exception.Message)"
        }

        # Per-user DEVMODE blob from the logged-on user's hive (resolved
        # above via Resolve-HkcuRoot). This is what the print dialog in
        # apps actually reads, distinct from the spooler-side default
        # captured by Get-PrintConfiguration.
        try {
            if (Test-Path $devModeKey) {
                $itemProp = Get-ItemProperty -Path $devModeKey -Name $p.Name -ErrorAction SilentlyContinue
                if ($null -ne $itemProp) {
                    $blob = $itemProp.$($p.Name)
                    if ($null -ne $blob -and $blob -is [byte[]] -and $blob.Length -gt 0) {
                        $b64 = [Convert]::ToBase64String($blob)
                        $b64 | Out-File -FilePath $devModePath -Encoding ASCII -Force -ErrorAction Stop
                        $devModeOk = $true
                    }
                }
            }
        }
        catch {
            $warnings += "Per-user DEVMODE capture failed for '$($p.Name)': $($_.Exception.Message)"
        }

        # Hardware/installable-options dump from HKLM. Vendor-specific
        # binary blobs that drive things like "Cassette 2/3 installed",
        # finisher attached, duplex unit, etc. Without this, Add-Printer
        # on the target re-initializes to driver minimum hardware config
        # and trays disappear from the print UI.
        try {
            $hwRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($p.Name)\PrinterDriverData"
            if (Test-Path $hwRegPath) {
                $key = Get-Item -Path $hwRegPath -ErrorAction Stop
                $hwDump = [ordered]@{}
                foreach ($vname in $key.GetValueNames()) {
                    $vtype = $key.GetValueKind($vname).ToString()
                    $vdata = $key.GetValue($vname)
                    $encoded = switch ($vtype) {
                        'Binary'       { if ($null -ne $vdata) { [Convert]::ToBase64String([byte[]]$vdata) } else { $null } }
                        'String'       { [string]$vdata }
                        'ExpandString' { [string]$vdata }
                        'MultiString'  { @($vdata) }
                        'DWord'        { [int]$vdata }
                        'QWord'        { [long]$vdata }
                        default        { "$vdata" }
                    }
                    $hwDump[$vname] = [ordered]@{
                        Type = $vtype
                        Data = $encoded
                    }
                }
                $hwDump | ConvertTo-Json -Depth 4 | Out-File -FilePath $hwConfigPath -Encoding UTF8 -Force
                $hwConfigOk = $true
            }
        }
        catch {
            $warnings += "Hardware config capture failed for '$($p.Name)': $($_.Exception.Message)"
        }

        $printSettingsRefs[$p.Name] = @{
            XmlFile      = if ($xmlOk)      { $xmlFile }      else { $null }
            PropFile     = if ($propOk)     { $propFile }     else { $null }
            DevModeFile  = if ($devModeOk)  { $devModeFile }  else { $null }
            HwConfigFile = if ($hwConfigOk) { $hwConfigFile } else { $null }
        }
    }

    Show-Success "Print settings captured for $($printSettingsRefs.Count) printer(s)"
}


# ---- 5.6: driver payload export (optional)
$exportedDrivers = @{}

if ($includeDriverBinaries -and $infPackagesToExport.Count -gt 0) {
    Show-Info "Exporting driver packages with pnputil..."
    $driversDir = Join-Path $backupDir "drivers"
    $null = New-Item -ItemType Directory -Path $driversDir -Force -ErrorAction SilentlyContinue

    foreach ($oemInf in @($infPackagesToExport.Keys)) {
        $outDir = Join-Path $driversDir $oemInf
        try {
            $null = New-Item -ItemType Directory -Path $outDir -Force -ErrorAction Stop
        }
        catch {
            Show-Error "Failed to create driver output dir: $outDir"
            $warnings += "Driver export skipped for ${oemInf}: cannot create output directory"
            $failCount++
            continue
        }

        Show-Info "  Exporting: $oemInf"
        $null = & pnputil /export-driver $oemInf $outDir 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0) {
            $size = (Get-ChildItem -Path $outDir -Recurse -File -ErrorAction SilentlyContinue |
                     Measure-Object -Property Length -Sum).Sum
            if ($null -eq $size) { $size = 0 }
            $exportedDrivers[$oemInf] = @{
                Folder    = "drivers/$oemInf"
                SizeBytes = [long]$size
            }
            $sizeMB = [math]::Round($size / 1MB, 1)
            Show-Success "    Exported $oemInf ($sizeMB MB)"
        }
        else {
            Show-Error "    pnputil failed for ${oemInf} (exit ${exitCode})"
            $warnings += "Driver export failed for ${oemInf}: pnputil exit code ${exitCode}"
            $failCount++
        }
    }
}

# Inbox-driver warning aggregation (informational)
foreach ($d in $driverInfoList) {
    if ($d.IsInboxDriver) {
        $warnings += "Driver '$($d.DriverName)': inbox driver, payload not exported (Windows-supplied)"
    }
}


# ---- 5.7: default printer
$defaultPrinter = $null
try {
    $defaultPrinter = (Get-CimInstance -ClassName Win32_Printer -Filter "Default=$true" -ErrorAction Stop |
                       Select-Object -First 1).Name
}
catch {
    $warnings += "Default printer detection failed: $($_.Exception.Message)"
}


# ---- 5.8: manifest.json
Show-Info "Writing manifest.json..."

$hwUid = $null
try {
    $hwUid = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop |
              Select-Object -First 1).UUID
}
catch {
    # hardwareUniqueId is informational only; leave null on failure
}

$osArch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    'arm64'
} elseif ([Environment]::Is64BitOperatingSystem) {
    'amd64'
} else {
    'x86'
}

$osVersion = [System.Environment]::OSVersion.Version.ToString()

$kernelVersionFile = Join-Path $PSScriptRoot "..\..\..\kernel\KERNEL_VERSION"
$kernelVersion = if (Test-Path $kernelVersionFile) {
    (Get-Content $kernelVersionFile -Raw).Trim()
} else {
    "unknown"
}

$moduleVersionFile = Join-Path $PSScriptRoot "VERSION"
$moduleVersion = if (Test-Path $moduleVersionFile) {
    (Get-Content $moduleVersionFile -Raw).Trim()
} else {
    "unknown"
}

$manifestPrinters = foreach ($p in $printers) {
    $ref = $printSettingsRefs[$p.Name]
    $matchedDriver = @($driverInfoList | Where-Object { $_.DriverName -eq $p.DriverName }) | Select-Object -First 1
    [PSCustomObject]@{
        name              = $p.Name
        driverName        = $p.DriverName
        portName          = $p.PortName
        shared            = [bool]$p.Shared
        shareName         = $p.ShareName
        comment           = $p.Comment
        location          = $p.Location
        published         = [bool]$p.Published
        isInboxDriver     = if ($matchedDriver) { [bool]$matchedDriver.IsInboxDriver } else { $false }
        printSettingsFile = if ($ref) { $ref.XmlFile      } else { $null }
        propertiesFile    = if ($ref) { $ref.PropFile     } else { $null }
        devModeFile       = if ($ref) { $ref.DevModeFile  } else { $null }
        hwConfigFile      = if ($ref) { $ref.HwConfigFile } else { $null }
    }
}

$manifestPorts = foreach ($p in $ports) {
    [PSCustomObject]@{
        name               = $p.Name
        portType           = Get-PortType -Port $p
        printerHostAddress = $p.PrinterHostAddress
        portNumber         = $p.PortNumber
        lprHostName        = $p.LprHostName
        lprQueueName       = $p.LprQueueName
        snmpEnabled        = [bool]$p.SnmpEnabled
        snmpCommunity      = $p.SnmpCommunity
        portMonitor        = $p.PortMonitor
    }
}

$manifestDrivers = foreach ($d in $driverInfoList) {
    $payload = if ($d.OemInf -and $exportedDrivers.ContainsKey($d.OemInf)) {
        $exportedDrivers[$d.OemInf]
    } else { $null }
    [PSCustomObject]@{
        driverName    = $d.DriverName
        manufacturer  = $d.Manufacturer
        driverVersion = $d.DriverVersion
        infBaseName   = $d.InfBaseName
        infOemFile    = $d.OemInf
        isInboxDriver = $d.IsInboxDriver
        backupFolder  = if ($payload) { $payload.Folder } else { $null }
    }
}

$driverBytes = 0
foreach ($v in $exportedDrivers.Values) {
    $driverBytes += $v.SizeBytes
}

$manifest = [ordered]@{
    schemaVersion       = 1
    manifestType        = "fabriq-printer-backup"
    backupVersion       = $moduleVersion
    fabriqKernelVersion = $kernelVersion
    collectedAt         = (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
    computerName        = $pcName
    hardwareUniqueId    = $hwUid
    osVersion           = $osVersion
    osArch              = $osArch
    defaultPrinter      = $defaultPrinter
    counts              = [ordered]@{
        printer          = @($manifestPrinters).Count
        port             = @($manifestPorts).Count
        driverRegistered = @($manifestDrivers).Count
        infPackage       = @($exportedDrivers.Keys).Count
    }
    sizes               = [ordered]@{
        totalBytes  = 0
        driverBytes = [long]$driverBytes
    }
    includes            = [ordered]@{
        driverBinaries = $includeDriverBinaries
        printSettings  = $includePrintSettings
    }
    items               = [ordered]@{
        printers = @($manifestPrinters)
        ports    = @($manifestPorts)
        drivers  = @($manifestDrivers)
    }
    warnings            = @($warnings)
}

$manifestPath = Join-Path $backupDir "manifest.json"
try {
    $manifest | ConvertTo-Json -Depth 8 | Out-File -FilePath $manifestPath -Encoding UTF8 -Force
    Show-Success "manifest.json written"
}
catch {
    Show-Error "Failed to write manifest.json: $_"
    $failCount++
}

# Re-compute total bytes AFTER manifest write (so the total includes the manifest itself),
# then patch the manifest.sizes.totalBytes field in place.
$totalBytes = (Get-ChildItem -Path $backupDir -Recurse -File -ErrorAction SilentlyContinue |
               Measure-Object -Property Length -Sum).Sum
if ($null -eq $totalBytes) { $totalBytes = 0 }
$manifest.sizes.totalBytes = [long]$totalBytes
try {
    $manifest | ConvertTo-Json -Depth 8 | Out-File -FilePath $manifestPath -Encoding UTF8 -Force
}
catch {
    Show-Warning "Failed to update totalBytes in manifest.json: $_"
}


# ---- 5.9: _restore_notes.txt (human-readable hints)
$inboxDriverNames = @($driverInfoList | Where-Object { $_.IsInboxDriver } | ForEach-Object { $_.DriverName })
$restoreNotes = @"
Printer Backup - Restore Notes
================================
Backup taken    : $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
Source PC       : $pcName$(if ($pcName -ne $env:COMPUTERNAME) { " (from hostlist OldPCname)" } else { "" })
Computer name   : $env:COMPUTERNAME$(if ($pcName -ne $env:COMPUTERNAME) { " (actual Windows hostname at backup time)" } else { "" })
OS / Arch       : $osVersion / $osArch

A companion restore module will read manifest.json from this folder.
Restore is only guaranteed on a target PC with the same OS version
and architecture as recorded above. Cross-architecture restore is
NOT supported.

Driver packages exported:
  $(if ($exportedDrivers.Count -gt 0) { @($exportedDrivers.Keys) -join ', ' } else { '(none)' })

Drivers NOT exported (inbox / Windows-supplied):
  $(if ($inboxDriverNames.Count -gt 0) { $inboxDriverNames -join ', ' } else { '(none)' })

Warnings:
$(if ($warnings.Count -gt 0) { ($warnings | ForEach-Object { "  - $_" }) -join "`n" } else { '  (none)' })
"@
$restoreNotes | Out-File -FilePath (Join-Path $backupDir "_restore_notes.txt") -Encoding UTF8 -Force

Write-Host ""


# ========================================
# Step 5.5: Post-Apply Verification
# ========================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Post-Apply Verification" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$verifyFail = 0

$requiredFiles = @(
    "manifest.json",
    "printers.json",
    "ports.json",
    "drivers_registered.json",
    "drivers_inf_inventory.json",
    "_restore_notes.txt"
)
foreach ($rf in $requiredFiles) {
    $path = Join-Path $backupDir $rf
    if (-not (Test-Path $path)) {
        Write-Host "  [VERIFY FAILED] $rf - missing" -ForegroundColor Red
        $verifyFail++
        continue
    }
    $sz = (Get-Item $path).Length
    if ($sz -eq 0) {
        Write-Host "  [VERIFY FAILED] $rf - zero bytes" -ForegroundColor Red
        $verifyFail++
        continue
    }
    Write-Host "  [VERIFIED] $rf ($sz bytes)" -ForegroundColor Green
}

# manifest.json parses as valid JSON with expected type/schema
try {
    $reread = Get-Content -Path (Join-Path $backupDir "manifest.json") -Raw | ConvertFrom-Json
    if ($reread.manifestType -eq "fabriq-printer-backup" -and $reread.schemaVersion -eq 1) {
        Write-Host "  [VERIFIED] manifest.json schemaVersion=1, manifestType=fabriq-printer-backup" -ForegroundColor Green
    }
    else {
        Write-Host "  [VERIFY FAILED] manifest.json - unexpected type/schemaVersion" -ForegroundColor Red
        $verifyFail++
    }
}
catch {
    Write-Host "  [VERIFY FAILED] manifest.json - parse error: $_" -ForegroundColor Red
    $verifyFail++
}

# drivers/oem*.inf folder count matches expected exports
if ($includeDriverBinaries) {
    $expected = $exportedDrivers.Count
    $actualDir = Join-Path $backupDir "drivers"
    $actual = if (Test-Path $actualDir) {
        @(Get-ChildItem -Path $actualDir -Directory -ErrorAction SilentlyContinue).Count
    } else { 0 }
    if ($actual -eq $expected) {
        Write-Host "  [VERIFIED] drivers/ contains $actual package(s)" -ForegroundColor Green
    } else {
        Write-Host "  [VERIFY FAILED] drivers/ count mismatch (expected $expected, actual $actual)" -ForegroundColor Red
        $verifyFail++
    }
}

# printsettings/*.xml count matches manifest expectations
if ($includePrintSettings) {
    $expected = @($manifestPrinters | Where-Object { $_.printSettingsFile }).Count
    $actualDir = Join-Path $backupDir "printsettings"
    $actual = if (Test-Path $actualDir) {
        @(Get-ChildItem -Path $actualDir -Filter *.xml -ErrorAction SilentlyContinue).Count
    } else { 0 }
    if ($actual -eq $expected) {
        Write-Host "  [VERIFIED] printsettings/ contains $actual XML file(s)" -ForegroundColor Green
    } else {
        Write-Host "  [VERIFY FAILED] printsettings/ count mismatch (expected $expected, actual $actual)" -ForegroundColor Red
        $verifyFail++
    }
}

Write-Host ""
$verified = ($verifyFail -eq 0)


# ========================================
# Step 6: Result Summary
# ========================================
$totalMB = [math]::Round($totalBytes / 1MB, 1)

Show-Separator
Write-Host "Printer Backup Results" -ForegroundColor Cyan
Show-Separator
Write-Host "  Location:  $backupDir" -ForegroundColor White
Write-Host "  Size:      $totalMB MB" -ForegroundColor White
Write-Host "  Printers:  $($printers.Count)" -ForegroundColor White
Write-Host "  Drivers:   $($printerDrivers.Count) registered, $($exportedDrivers.Count) package(s) exported" -ForegroundColor White
if ($warnings.Count -gt 0) {
    Write-Host "  Warnings:  $($warnings.Count) (see manifest.json / _restore_notes.txt)" -ForegroundColor Yellow
}
if ($verified) {
    Write-Host "  Verified:  PASS" -ForegroundColor Green
} else {
    Write-Host "  Verified:  FAIL" -ForegroundColor Red
}
Show-Separator
Write-Host ""

$finalStatus = if ($failCount -eq 0 -and $verified) {
    "Success"
} elseif ($failCount -gt 0) {
    "Partial"
} else {
    "Partial"
}

return (New-ModuleResult `
    -Status $finalStatus `
    -Message "Backup written to $backupDir ($totalMB MB)" `
    -Details @($warnings) `
    -Verified $verified)
