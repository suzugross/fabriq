# ========================================
# Temporary IP Address Assignment Script
# ========================================
# [PURPOSE]
# Show the operator a GUI dialog listing pool candidates from CSV; the
# operator picks one and the script assigns it to the target NIC. Used
# when the production IP is still occupied by the old PC being replaced.
#
# [WHY GUI INSTEAD OF AUTOMATIC PROBE]
# Pre-probing the pool (ICMP/ARP) is unreliable in temp-IP scenarios:
#   - Brand-new kitting PCs may have no IPv4 address yet -> cannot ARP
#   - Cannot detect a previously assigned IP if its target PC is offline
#     (LAN-disconnected for transport, etc.)
# The honest design is: operator picks based on team coordination, then
# Windows DAD (Duplicate Address Detection) catches collisions at the
# moment of assignment. If DAD fires, the dialog reopens with the bad
# IP grayed out for re-selection.
#
# [WORKFLOW]
#   1. Validate CSV rows
#   2. Resolve NIC by AdapterPattern (single Up Ethernet)
#   3. Subnet sanity warning (informational only)
#   4. Show selection dialog (modal, regardless of AutoPilot)
#   5. Assign + DAD verify; on Duplicate, reopen dialog with exclusion
#   6. Post-Apply Verification (AddressState, route, DNS, gateway ping)
# ========================================

Write-Host ""
Show-Separator
Write-Host "Temporary IP Address Assignment" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Helper: Test-IPv4Address (strict dotted-quad)
# ========================================
function Test-IPv4Address {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $trimmed = $Value.Trim()
    if ($trimmed -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { return $false }
    $parsed = $null
    if (-not [System.Net.IPAddress]::TryParse($trimmed, [ref]$parsed)) { return $false }
    return ($parsed.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork)
}

# ========================================
# Helper: Test-RowValidity
# ========================================
function Test-RowValidity {
    param($Row)
    $issues = @()

    if (-not (Test-IPv4Address -Value $Row.IPAddress)) {
        $issues += "IPAddress is not a valid IPv4 address: '$($Row.IPAddress)'"
    }
    $prefix = 0
    if (-not [int]::TryParse($Row.SubnetPrefix, [ref]$prefix) -or $prefix -lt 1 -or $prefix -gt 32) {
        $issues += "SubnetPrefix must be 1-32 (got: '$($Row.SubnetPrefix)')"
    }
    if ($Row.Gateway -and -not (Test-IPv4Address -Value $Row.Gateway)) {
        $issues += "Gateway is not a valid IPv4 address: '$($Row.Gateway)'"
    }
    foreach ($dnsCol in @('DNS1', 'DNS2', 'DNS3')) {
        $dns = $Row.$dnsCol
        if ($dns -and -not (Test-IPv4Address -Value $dns)) {
            $issues += "$dnsCol is not a valid IPv4 address: '$dns'"
        }
    }
    if ([string]::IsNullOrWhiteSpace($Row.AdapterPattern)) {
        $issues += "AdapterPattern is empty"
    }
    return $issues
}

# ========================================
# Helper: Resolve-TargetAdapter
# ========================================
function Resolve-TargetAdapter {
    param([string]$Pattern)

    # Avoid the automatic variable $matches (set by -match operator)
    $matchedAdapters = @(Get-NetAdapter -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like $Pattern -and
            $_.Status -eq 'Up' -and
            $_.Virtual -eq $false
        } | Sort-Object ifIndex)

    if ($matchedAdapters.Count -eq 0) {
        Show-Error "No Up non-virtual NIC matched pattern '$Pattern'"
        $allUp = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' })
        if ($allUp.Count -gt 0) {
            Show-Info "Available Up adapters:"
            foreach ($a in $allUp) {
                Write-Host "  - $($a.Name) (ifIndex=$($a.ifIndex), Virtual=$($a.Virtual))" -ForegroundColor DarkGray
            }
        }
        return $null
    }
    if ($matchedAdapters.Count -gt 1) {
        Show-Error "Pattern '$Pattern' matches $($matchedAdapters.Count) NICs (must be unique):"
        foreach ($a in $matchedAdapters) {
            Write-Host "  - $($a.Name) (ifIndex=$($a.ifIndex))" -ForegroundColor DarkGray
        }
        return $null
    }
    return $matchedAdapters[0]
}

# ========================================
# Helper: Convert-IPToNetwork
# ========================================
function Convert-IPToNetwork {
    param([string]$IPAddress, [int]$Prefix)
    $ipBytes  = ([System.Net.IPAddress]::Parse($IPAddress)).GetAddressBytes()
    $netBytes = New-Object byte[] 4
    for ($i = 0; $i -lt 4; $i++) {
        $bitsInThisByte = [Math]::Min(8, [Math]::Max(0, $Prefix - ($i * 8)))
        if ($bitsInThisByte -eq 0) {
            $maskByte = 0
        }
        else {
            $maskByte = [byte](256 - [Math]::Pow(2, 8 - $bitsInThisByte))
        }
        $netBytes[$i] = $ipBytes[$i] -band $maskByte
    }
    return ([System.Net.IPAddress]::new($netBytes)).IPAddressToString
}

# ========================================
# Helper: Get-AdapterSubnetCheck
# ========================================
function Get-AdapterSubnetCheck {
    param([int]$InterfaceIndex, [string]$PoolIPSample, [int]$PoolPrefix)

    $currentIps = @(Get-NetIPAddress -InterfaceIndex $InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '0.0.0.0' })

    if ($currentIps.Count -eq 0) {
        return @{ SameSubnet = $false; Reason = 'NIC has no usable IPv4 address' }
    }

    $poolNet = Convert-IPToNetwork -IPAddress $PoolIPSample -Prefix $PoolPrefix
    foreach ($ip in $currentIps) {
        $curNet = Convert-IPToNetwork -IPAddress $ip.IPAddress -Prefix $PoolPrefix
        if ($curNet -eq $poolNet) {
            return @{ SameSubnet = $true; Reason = "NIC IP $($ip.IPAddress) is in pool subnet $poolNet/$PoolPrefix" }
        }
    }
    return @{ SameSubnet = $false; Reason = "NIC IPs ($(($currentIps.IPAddress) -join ', ')) are not in pool subnet" }
}

# ========================================
# Helper: Set-TempIPAndVerify
# ========================================
# Apply IP / Gateway / DNS, wait for DAD, return verification result.
# Returns @{ Status='OK'/'Duplicate'/'Error'; Reason=<string>; Address=<obj> }.
function Set-TempIPAndVerify {
    param(
        [int]$InterfaceIndex,
        [string]$IPAddress,
        [int]$PrefixLength,
        [string]$Gateway,
        [string[]]$DnsServers
    )

    # Clean slate: remove existing IP / default route on this NIC
    try {
        Get-NetIPAddress -InterfaceIndex $InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.PrefixOrigin -ne 'WellKnown' } |
            Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue
    } catch { }
    try {
        Get-NetRoute -InterfaceIndex $InterfaceIndex -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
            Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue
    } catch { }

    # Assign (return value discarded; we re-fetch for verification below)
    try {
        $newIpParams = @{
            InterfaceIndex = $InterfaceIndex
            IPAddress      = $IPAddress
            PrefixLength   = $PrefixLength
            ErrorAction    = 'Stop'
        }
        if ($Gateway) { $newIpParams.DefaultGateway = $Gateway }
        New-NetIPAddress @newIpParams | Out-Null
    }
    catch {
        return @{ Status = 'Error'; Reason = "New-NetIPAddress failed: $_"; Address = $null }
    }

    # DNS (best effort)
    $dnsArray = @($DnsServers | Where-Object { $_ })
    if ($dnsArray.Count -gt 0) {
        try {
            Set-DnsClientServerAddress -InterfaceIndex $InterfaceIndex -ServerAddresses $dnsArray -ErrorAction Stop
        }
        catch {
            Show-Warning "Set-DnsClientServerAddress failed (continuing): $_"
        }
    }

    # Wait for DAD
    Start-Sleep -Milliseconds 500
    $check = Get-NetIPAddress -InterfaceIndex $InterfaceIndex -IPAddress $IPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue
    if (-not $check) {
        return @{ Status = 'Error'; Reason = 'Address disappeared after assignment'; Address = $null }
    }
    if ($check.AddressState -eq 'Tentative') {
        Start-Sleep -Milliseconds 1000
        $check = Get-NetIPAddress -InterfaceIndex $InterfaceIndex -IPAddress $IPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue
    }
    if (-not $check) {
        return @{ Status = 'Error'; Reason = 'Address disappeared during DAD wait'; Address = $null }
    }
    if ($check.AddressState -eq 'Duplicate') {
        try {
            Remove-NetIPAddress -InterfaceIndex $InterfaceIndex -IPAddress $IPAddress -Confirm:$false -ErrorAction SilentlyContinue
        } catch { }
        return @{ Status = 'Duplicate'; Reason = 'DAD detected duplicate address on the network'; Address = $null }
    }
    if ($check.AddressState -ne 'Preferred') {
        return @{ Status = 'Error'; Reason = "AddressState is $($check.AddressState) (expected Preferred)"; Address = $check }
    }
    return @{ Status = 'OK'; Reason = ''; Address = $check }
}

# ========================================
# Helper: Show-IPSelectionDialog
# ========================================
# Modal Windows.Forms dialog. Returns the selected CSV row (PSCustomObject)
# or $null if the operator cancelled. Excluded IPs (typically previously
# DAD-detected duplicates) are shown grayed out and cannot be picked.
function Show-IPSelectionDialog {
    param(
        [array]$PoolRows,
        [string]$NicName,
        [string]$CurrentIP,
        [string]$SubnetWarning,
        [string[]]$ExcludedIPs = @()
    )

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $form = New-Object System.Windows.Forms.Form
    $form.Text             = "Temp IP Address Selection"
    $form.Size             = New-Object System.Drawing.Size(720, 540)
    $form.StartPosition    = 'CenterScreen'
    $form.FormBorderStyle  = 'FixedDialog'
    $form.MaximizeBox      = $false
    $form.MinimizeBox      = $false
    $form.TopMost          = $true

    # Header
    $headerText = "Target NIC : $NicName`r`nCurrent IP : $(if ($CurrentIP) { $CurrentIP } else { '(none)' })"
    if ($SubnetWarning) {
        $headerText += "`r`n[!] $SubnetWarning"
    }
    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Location = New-Object System.Drawing.Point(15, 15)
    $headerLabel.Size     = New-Object System.Drawing.Size(680, 65)
    $headerLabel.Text     = $headerText
    $headerLabel.Font     = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($headerLabel)

    # Instructions
    $instrLabel = New-Object System.Windows.Forms.Label
    $instrLabel.Location  = New-Object System.Drawing.Point(15, 90)
    $instrLabel.Size      = New-Object System.Drawing.Size(680, 20)
    $instrLabel.Text      = "Pool (no probing performed - coordinate with team to avoid clash):"
    $instrLabel.Font      = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($instrLabel)

    # ListView
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location      = New-Object System.Drawing.Point(15, 115)
    $listView.Size          = New-Object System.Drawing.Size(680, 280)
    $listView.View          = 'Details'
    $listView.FullRowSelect = $true
    $listView.MultiSelect   = $false
    $listView.GridLines     = $true
    $listView.HideSelection = $false
    [void]$listView.Columns.Add("IP Address",  140)
    [void]$listView.Columns.Add("Prefix",       60)
    [void]$listView.Columns.Add("Gateway",     130)
    [void]$listView.Columns.Add("Description", 240)
    [void]$listView.Columns.Add("Status",      100)

    foreach ($row in $PoolRows) {
        $item = New-Object System.Windows.Forms.ListViewItem($row.IPAddress)
        [void]$item.SubItems.Add("/$($row.SubnetPrefix)")
        [void]$item.SubItems.Add($(if ($row.Gateway) { $row.Gateway } else { "" }))
        [void]$item.SubItems.Add($(if ($row.Description) { $row.Description } else { "" }))

        if ($row.IPAddress -in $ExcludedIPs) {
            [void]$item.SubItems.Add("[DUPLICATE]")
            $item.ForeColor = [System.Drawing.Color]::Gray
            $item.Tag       = "EXCLUDED"
        }
        else {
            if ($row.IPAddress -eq $CurrentIP) {
                [void]$item.SubItems.Add("[CURRENT]")
            }
            else {
                [void]$item.SubItems.Add("")
            }
            $item.Tag = $row
        }
        [void]$listView.Items.Add($item)
    }

    # Pre-select: prefer CURRENT, then first non-excluded
    $preselect = $null
    foreach ($it in $listView.Items) {
        if ($it.Tag -ne "EXCLUDED" -and $it.SubItems[4].Text -eq "[CURRENT]") {
            $preselect = $it; break
        }
    }
    if (-not $preselect) {
        foreach ($it in $listView.Items) {
            if ($it.Tag -ne "EXCLUDED") { $preselect = $it; break }
        }
    }
    if ($preselect) {
        $preselect.Selected = $true
        $preselect.EnsureVisible()
    }

    # Block selection of excluded items
    $listView.Add_ItemSelectionChanged({
        param($s, $e)
        if ($e.IsSelected -and $e.Item.Tag -eq "EXCLUDED") {
            $e.Item.Selected = $false
        }
    })
    $form.Controls.Add($listView)

    # Note label
    $noteLabel = New-Object System.Windows.Forms.Label
    $noteLabel.Location  = New-Object System.Drawing.Point(15, 405)
    $noteLabel.Size      = New-Object System.Drawing.Size(680, 40)
    $noteLabel.Text      = "Windows DAD will detect collisions visible on the LAN at assignment time. " +
                           "If a picked IP fires DAD, this dialog reopens with that IP grayed out."
    $noteLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $form.Controls.Add($noteLabel)

    # Buttons
    $assignBtn = New-Object System.Windows.Forms.Button
    $assignBtn.Location     = New-Object System.Drawing.Point(460, 455)
    $assignBtn.Size         = New-Object System.Drawing.Size(110, 32)
    $assignBtn.Text         = "Assign"
    $assignBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton      = $assignBtn
    $form.Controls.Add($assignBtn)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Location     = New-Object System.Drawing.Point(580, 455)
    $cancelBtn.Size         = New-Object System.Drawing.Size(110, 32)
    $cancelBtn.Text         = "Cancel"
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton      = $cancelBtn
    $form.Controls.Add($cancelBtn)

    $result = $form.ShowDialog()
    $picked = $null
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $listView.SelectedItems.Count -gt 0) {
        $sel = $listView.SelectedItems[0]
        if ($sel.Tag -is [PSCustomObject]) {
            $picked = $sel.Tag
        }
    }
    $form.Dispose()
    return $picked
}

# ========================================
# Step 0: Privilege Check
# ========================================
if (-not (Test-AdminPrivilege)) {
    Show-Error "Administrator privileges are required."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Administrator privileges required")
}

# ========================================
# Step 1: Load CSV
# ========================================
$csvPath = Join-Path $PSScriptRoot "temp_ipaddress_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "IPAddress", "SubnetPrefix", "AdapterPattern")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load temp_ipaddress_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}

# ========================================
# Step 2: Per-row validation
# ========================================
$validRows = @()
$invalidCount = 0
foreach ($row in $enabledItems) {
    $issues = Test-RowValidity -Row $row
    if ($issues.Count -eq 0) {
        $validRows += $row
    }
    else {
        Show-Error "[INVALID] $($row.IPAddress) : $($issues -join '; ')"
        $invalidCount++
    }
}
if ($validRows.Count -eq 0) {
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "All $invalidCount enabled rows are invalid")
}
if ($invalidCount -gt 0) {
    Show-Warning "$invalidCount invalid row(s) excluded from pool"
    Write-Host ""
}

# ========================================
# Step 3: Resolve target NIC
# ========================================
$patterns = @($validRows | ForEach-Object { $_.AdapterPattern.Trim() } | Sort-Object -Unique)
if ($patterns.Count -gt 1) {
    Show-Error "Mixed AdapterPatterns in pool: $($patterns -join ', ')"
    Show-Error "v1 supports only a single NIC per run. Split via Segment or separate CSVs."
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Mixed AdapterPatterns ($($patterns.Count))")
}
$pattern = $patterns[0]
Show-Info "Resolving NIC for pattern: $pattern"
$nic = Resolve-TargetAdapter -Pattern $pattern
if ($null -eq $nic) {
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to resolve NIC for pattern '$pattern'")
}
Show-Success "Target NIC: $($nic.Name) (ifIndex=$($nic.ifIndex), MAC=$($nic.MacAddress))"
Write-Host ""

# ========================================
# Step 3.5: Subnet sanity check (informational)
# ========================================
$samplePool = $validRows[0]
$subnetCheck = Get-AdapterSubnetCheck -InterfaceIndex $nic.ifIndex `
    -PoolIPSample $samplePool.IPAddress -PoolPrefix ([int]$samplePool.SubnetPrefix)
$subnetWarning = $null
if ($subnetCheck.SameSubnet) {
    Show-Info $subnetCheck.Reason
}
else {
    Show-Warning $subnetCheck.Reason
    $subnetWarning = $subnetCheck.Reason
}
Write-Host ""

# ========================================
# Step 4: GUI selection + Assignment loop (DAD retry)
# ========================================
# Surface a quick reference: current Preferred IP on this NIC (for sticky default)
$currentPreferredIP = $null
$currentIPs = @(Get-NetIPAddress -InterfaceIndex $nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object { $_.PrefixOrigin -ne 'WellKnown' -and $_.AddressState -eq 'Preferred' })
if ($currentIPs.Count -gt 0) {
    $currentPreferredIP = $currentIPs[0].IPAddress
}

$excludedIPs = @()
$selected    = $null
$attempt     = 0

while ($true) {
    $attempt++
    Show-Info "Awaiting operator selection (attempt $attempt)..."

    $picked = Show-IPSelectionDialog `
        -PoolRows      $validRows `
        -NicName       $nic.Name `
        -CurrentIP     $currentPreferredIP `
        -SubnetWarning $subnetWarning `
        -ExcludedIPs   $excludedIPs

    if ($null -eq $picked) {
        Write-Host ""
        Show-Info "Cancelled by operator."
        return (New-ModuleResult -Status "Cancelled" -Message "Operator cancelled the selection dialog")
    }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Selected by operator: $($picked.IPAddress)/$($picked.SubnetPrefix)" -ForegroundColor Yellow
    Write-Host "  Gateway     : $(if ($picked.Gateway) { $picked.Gateway } else { '(not set)' })" -ForegroundColor White
    $dnsList = @($picked.DNS1, $picked.DNS2, $picked.DNS3 | Where-Object { $_ })
    Write-Host "  DNS         : $(if ($dnsList) { $dnsList -join ', ' } else { '(not changed)' })" -ForegroundColor White
    Write-Host "  Description : $($picked.Description)" -ForegroundColor DarkGray
    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host ""

    # Sticky shortcut: if operator picked the IP that's already on the NIC, skip reassignment
    if ($picked.IPAddress -eq $currentPreferredIP) {
        Show-Info "Sticky: operator picked the current NIC IP. No reassignment needed."
        Write-Host ""
        $selected = $picked
        break
    }

    # Apply
    $result = Set-TempIPAndVerify -InterfaceIndex $nic.ifIndex `
        -IPAddress    $picked.IPAddress `
        -PrefixLength ([int]$picked.SubnetPrefix) `
        -Gateway      $picked.Gateway `
        -DnsServers   @($picked.DNS1, $picked.DNS2, $picked.DNS3)

    switch ($result.Status) {
        'OK' {
            Show-Success "Assigned $($picked.IPAddress)/$($picked.SubnetPrefix) on $($nic.Name) (AddressState=Preferred)"
            $selected = $picked
        }
        'Duplicate' {
            Show-Warning "[DAD] $($picked.IPAddress) detected as duplicate on the network."
            Show-Info    "Reopening dialog with this IP excluded. Pick another."
            $excludedIPs += $picked.IPAddress
            Write-Host ""
            continue
        }
        default {
            Show-Error "Assignment error for $($picked.IPAddress): $($result.Reason)"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Assignment failed: $($result.Reason)")
        }
    }
    break
}

# ========================================
# Step 5: Post-Apply Verification
# ========================================
Show-Info "Verifying assignment..."
Write-Host ""
$verifyIssues = @()

$verifyIp = Get-NetIPAddress -InterfaceIndex $nic.ifIndex -IPAddress $selected.IPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue
if (-not $verifyIp) {
    $verifyIssues += "IP not present after assignment"
}
else {
    if ($verifyIp.PrefixLength -ne [int]$selected.SubnetPrefix) {
        $verifyIssues += "PrefixLength mismatch (expected=$($selected.SubnetPrefix), actual=$($verifyIp.PrefixLength))"
    }
    if ($verifyIp.AddressState -ne 'Preferred') {
        $verifyIssues += "AddressState=$($verifyIp.AddressState) (expected Preferred)"
    }
}

if ($selected.Gateway) {
    $route = Get-NetRoute -InterfaceIndex $nic.ifIndex -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
        Where-Object { $_.NextHop -eq $selected.Gateway }
    if (-not $route) {
        $verifyIssues += "Default route via $($selected.Gateway) not found"
    }
    $gwReachable = $false
    try {
        $gwReachable = Test-Connection -ComputerName $selected.Gateway -Quiet -Count 2 -TimeoutSeconds 1 -ErrorAction SilentlyContinue
    } catch { }
    if ($gwReachable) {
        Write-Host "  [L3 OK] Gateway $($selected.Gateway) responds to ping" -ForegroundColor Green
    }
    else {
        Show-Warning "  Gateway $($selected.Gateway) does not respond to ping (may be ICMP-blocked)"
    }
}

$dnsExpected = @($selected.DNS1, $selected.DNS2, $selected.DNS3 | Where-Object { $_ })
if ($dnsExpected.Count -gt 0) {
    $dnsActual = @((Get-DnsClientServerAddress -InterfaceIndex $nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses)
    foreach ($d in $dnsExpected) {
        if ($d -notin $dnsActual) {
            $verifyIssues += "DNS '$d' not present in actual list ($($dnsActual -join ', '))"
        }
    }
}

if ($verifyIssues.Count -eq 0) {
    Write-Host "  [VERIFIED] $($selected.IPAddress)/$($selected.SubnetPrefix) on $($nic.Name)" -ForegroundColor Green
    $verified = $true
}
else {
    foreach ($v in $verifyIssues) { Write-Host "  [VERIFY FAILED] $v" -ForegroundColor Red }
    $verified = $false
}
Write-Host ""

# ========================================
# Step 6: Result
# ========================================
$msg = "Assigned $($selected.IPAddress)/$($selected.SubnetPrefix) on $($nic.Name)"
if ($excludedIPs.Count -gt 0) {
    $msg += " (DAD-excluded during selection: $($excludedIPs -join ', '))"
}
return (New-ModuleResult -Status "Success" -Message $msg -Verified $verified)
