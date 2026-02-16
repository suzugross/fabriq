#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Power Option Configuration Script
.DESCRIPTION
    Reads power setting profiles from a CSV file and applies the selected profile.
.NOTES
    Administrator privileges are required.
#>

# ========================================
# Global Variables
# ========================================
$script:CsvPath = Join-Path $PSScriptRoot "power_list.csv"

# Power Plan GUIDs
$script:PowerPlanGuids = @{
    'BALANCED'          = '381b4222-f694-41f0-9685-ff5bb260df2e'
    'HIGH_PERFORMANCE'  = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
    'POWER_SAVER'       = 'a1841308-3541-4fab-bc81-f71556f20b4a'
}

# Button Action Value Mapping
$script:ActionValues = @{
    'NOTHING'   = 0
    'SLEEP'     = 1
    'HIBERNATE' = 2
    'SHUTDOWN'  = 3
}

# Power Setting GUIDs (for powercfg)
$script:SettingGuids = @{
    'PowerButton'       = '7648efa3-dd9c-4e3e-b566-50f929386280'
    'SleepButton'       = '96996bc0-ad50-47ec-923b-6f41874dd9eb'
    'LidClose'          = '5ca83367-6e45-459f-a27b-476b1d01c936'
    'ProcessorMinState' = '893dee8e-2bef-41e0-89c6-b55d0929964c'
    'ProcessorMaxState' = 'bc5038f7-23e0-4960-96da-33abaf5935ec'
}

# Subgroup GUIDs
$script:SubGroupGuids = @{
    'PowerButtons' = '4f971e89-eebd-4455-a8de-9e59040e7347'
    'Processor'    = '54533251-82be-4824-96c1-47b60b740d00'
}

# Timeout Setting GUIDs (for idempotency checks)
$script:TimeoutGuids = @{
    'monitor-ac'   = @{ SubGroup = '7516b95f-f776-4464-8c53-06167f40cc99'; Setting = '3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e' }
    'monitor-dc'   = @{ SubGroup = '7516b95f-f776-4464-8c53-06167f40cc99'; Setting = '3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e' }
    'standby-ac'   = @{ SubGroup = '238c9fa8-0aad-41ed-83f4-97be242c8f20'; Setting = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da' }
    'standby-dc'   = @{ SubGroup = '238c9fa8-0aad-41ed-83f4-97be242c8f20'; Setting = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da' }
    'hibernate-ac' = @{ SubGroup = '238c9fa8-0aad-41ed-83f4-97be242c8f20'; Setting = '9d7815a6-7ee4-497e-8888-515a05f02364' }
    'hibernate-dc' = @{ SubGroup = '238c9fa8-0aad-41ed-83f4-97be242c8f20'; Setting = '9d7815a6-7ee4-497e-8888-515a05f02364' }
    'disk-ac'      = @{ SubGroup = '0012ee47-9041-4b5d-9b77-535fba8b1442'; Setting = '6738e2c4-e8a5-4a42-b16a-e040e769756e' }
    'disk-dc'      = @{ SubGroup = '0012ee47-9041-4b5d-9b77-535fba8b1442'; Setting = '6738e2c4-e8a5-4a42-b16a-e040e769756e' }
}

# Idempotency counters
$script:SkipCount = 0
$script:ChangeCount = 0

# ========================================
# Initialization Function
# ========================================
function Initialize-Script {
    Write-Host "Starting Power Option Configuration Script`n" -ForegroundColor Cyan
    
    # Check CSV File Existence
    if (-not (Test-Path $script:CsvPath)) {
        Write-Host "Error: CSV file not found: $script:CsvPath" -ForegroundColor Red
        throw "CSV file does not exist"
    }
    
    Write-Host "CSV file confirmed: $script:CsvPath`n" -ForegroundColor Green
}

# ========================================
# CSV Import Function
# ========================================
function Import-PowerSettingsCsv {
    try {
        Write-Host "Loading CSV file..." -ForegroundColor Gray
        $csvData = Import-Csv -Path $script:CsvPath -Encoding Default
        
        if ($csvData.Count -eq 0) {
            throw "No data in CSV file"
        }
        
        Write-Host "Loaded $($csvData.Count) profiles`n" -ForegroundColor Green
        return $csvData
    }
    catch {
        Write-Host "Error: Failed to load CSV - $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# ========================================
# Value Conversion Function
# ========================================
function ConvertTo-SettingValue {
    param(
        [string]$Value
    )
    
    # Skip if empty or "-"
    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq '-') {
        return $null
    }
    
    return $Value
}

# ========================================
# Get Current Power Setting Value (for idempotency)
# ========================================
function Get-PowerConfigValue {
    param(
        [string]$PlanGuid,
        [string]$SubGroupGuid,
        [string]$SettingGuid,
        [ValidateSet('AC', 'DC')]
        [string]$PowerSource
    )

    try {
        $output = & powercfg /QUERY $PlanGuid $SubGroupGuid $SettingGuid 2>&1
        if ($LASTEXITCODE -ne 0) { return $null }

        $pattern = if ($PowerSource -eq 'AC') {
            'Current AC Power Setting Index:\s+0x([0-9a-fA-F]+)'
        } else {
            'Current DC Power Setting Index:\s+0x([0-9a-fA-F]+)'
        }

        foreach ($line in $output) {
            if ($line -match $pattern) {
                return [Convert]::ToInt64($matches[1], 16)
            }
        }
    }
    catch { }
    return $null
}

# ========================================
# Get Current Timeout Value in Minutes
# ========================================
function Get-TimeoutValue {
    param(
        [string]$TimeoutType,
        [string]$PlanGuid
    )

    $guidInfo = $script:TimeoutGuids[$TimeoutType]
    if (-not $guidInfo) { return $null }

    $source = if ($TimeoutType -match '-ac$') { 'AC' } else { 'DC' }
    $seconds = Get-PowerConfigValue -PlanGuid $PlanGuid -SubGroupGuid $guidInfo.SubGroup `
                -SettingGuid $guidInfo.Setting -PowerSource $source

    if ($null -eq $seconds) { return $null }

    # powercfg /QUERY returns seconds, /CHANGE expects minutes
    return [math]::Floor($seconds / 60)
}

# ========================================
# Menu Display Function
# ========================================
function Show-ProfileMenu {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Profiles
    )
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Select Power Option Profile" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $Profiles.Count; $i++) {
        $profile = $Profiles[$i]
        Write-Host "[$($i + 1)] " -NoNewline -ForegroundColor Yellow
        Write-Host "$($profile.ProfileName)" -NoNewline -ForegroundColor White
        Write-Host " - $($profile.Description)" -ForegroundColor Gray
    }
    
    Write-Host "`n[0] Exit" -ForegroundColor Red
    Write-Host ""
    
    do {
        $selection = Read-Host "Please select (0-$($Profiles.Count))"
        $selectionNum = $null
        $validInput = [int]::TryParse($selection, [ref]$selectionNum)
    } while (-not $validInput -or $selectionNum -lt 0 -or $selectionNum -gt $Profiles.Count)
    
    if ($selectionNum -eq 0) {
        return $null
    }
    
    $selectedProfile = $Profiles[$selectionNum - 1]
    Write-Host "`nProfile '$($selectedProfile.ProfileName)' selected`n" -ForegroundColor Green
    
    return $selectedProfile
}

# ========================================
# Settings Confirmation Function
# ========================================
function Confirm-ApplySettings {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile
    )
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Settings to Apply" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Profile Name: " -NoNewline
    Write-Host $Profile.ProfileName -ForegroundColor Yellow
    Write-Host "Description: " -NoNewline
    Write-Host $Profile.Description -ForegroundColor Gray
    Write-Host "Power Plan: " -NoNewline
    Write-Host $Profile.PowerPlan -ForegroundColor Green
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    while ($true) {
        $confirmation = Read-Host "Apply these settings? (Y/N)"
        if ($confirmation -eq 'Y' -or $confirmation -eq 'y') { return $true }
        if ($confirmation -eq 'N' -or $confirmation -eq 'n') { return $false }
        Write-Host "[INFO] Please enter Y or N" -ForegroundColor Yellow
    }
}

# ========================================
# Power Plan Setting Function
# ========================================
function Set-PowerPlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PlanName
    )

    try {
        $planGuid = $script:PowerPlanGuids[$PlanName]

        if (-not $planGuid) {
            Write-Host "Warning: Unknown power plan: $PlanName" -ForegroundColor Yellow
            return $false
        }

        # Idempotency check
        $currentGuid = Get-ActivePowerPlanGuid
        if ($currentGuid -eq $planGuid) {
            Write-Host "[SKIP] Power plan already '$PlanName'" -ForegroundColor Gray
            $script:SkipCount++
            return $true
        }

        Write-Host "Changing power plan to '$PlanName'..." -ForegroundColor Gray

        $result = & powercfg /S $planGuid 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Changed power plan to '$PlanName'" -ForegroundColor Green
            $script:ChangeCount++
            return $true
        }
        else {
            Write-Host "[ERROR] Failed to change power plan: $result" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Error: Failed to set power plan - $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# ========================================
# Get Current Power Plan GUID
# ========================================
function Get-ActivePowerPlanGuid {
    try {
        $output = & powercfg /GETACTIVESCHEME
        if ($output -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
            return $matches[1]
        }
    }
    catch {
        Write-Host "Warning: Failed to get current power plan" -ForegroundColor Yellow
    }
    return $null
}

# ========================================
# Display Settings Function
# ========================================
function Set-DisplaySettings {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile,
        [string]$PlanGuid
    )

    # On AC Power
    $acValue = ConvertTo-SettingValue $Profile.Display_TurnOff_AC
    if ($null -ne $acValue) {
        $current = Get-TimeoutValue -TimeoutType 'monitor-ac' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$acValue) {
            Write-Host "[SKIP] Display Turn Off (AC): already ${acValue} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE monitor-timeout-ac $acValue | Out-Null
                Write-Host "[OK] Display Turn Off (AC): ${acValue} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set display (AC)" -ForegroundColor Yellow
            }
        }
    }

    # On Battery
    $batteryValue = ConvertTo-SettingValue $Profile.Display_TurnOff_Battery
    if ($null -ne $batteryValue) {
        $current = Get-TimeoutValue -TimeoutType 'monitor-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$batteryValue) {
            Write-Host "[SKIP] Display Turn Off (Battery): already ${batteryValue} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE monitor-timeout-dc $batteryValue | Out-Null
                Write-Host "[OK] Display Turn Off (Battery): ${batteryValue} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set display (Battery)" -ForegroundColor Yellow
            }
        }
    }
}

# ========================================
# Sleep Settings Function
# ========================================
function Set-SleepSettings {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile,
        [string]$PlanGuid
    )

    # Sleep - AC Power
    $sleepAc = ConvertTo-SettingValue $Profile.Sleep_After_AC
    if ($null -ne $sleepAc) {
        $current = Get-TimeoutValue -TimeoutType 'standby-ac' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$sleepAc) {
            Write-Host "[SKIP] Sleep After (AC): already ${sleepAc} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE standby-timeout-ac $sleepAc | Out-Null
                Write-Host "[OK] Sleep After (AC): ${sleepAc} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set sleep (AC)" -ForegroundColor Yellow
            }
        }
    }

    # Sleep - Battery
    $sleepDc = ConvertTo-SettingValue $Profile.Sleep_After_Battery
    if ($null -ne $sleepDc) {
        $current = Get-TimeoutValue -TimeoutType 'standby-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$sleepDc) {
            Write-Host "[SKIP] Sleep After (Battery): already ${sleepDc} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE standby-timeout-dc $sleepDc | Out-Null
                Write-Host "[OK] Sleep After (Battery): ${sleepDc} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set sleep (Battery)" -ForegroundColor Yellow
            }
        }
    }

    # Hibernate - AC Power
    $hibernateAc = ConvertTo-SettingValue $Profile.Hibernate_After_AC
    if ($null -ne $hibernateAc) {
        $current = Get-TimeoutValue -TimeoutType 'hibernate-ac' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$hibernateAc) {
            Write-Host "[SKIP] Hibernate After (AC): already ${hibernateAc} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE hibernate-timeout-ac $hibernateAc | Out-Null
                Write-Host "[OK] Hibernate After (AC): ${hibernateAc} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set hibernate (AC)" -ForegroundColor Yellow
            }
        }
    }

    # Hibernate - Battery
    $hibernateDc = ConvertTo-SettingValue $Profile.Hibernate_After_Battery
    if ($null -ne $hibernateDc) {
        $current = Get-TimeoutValue -TimeoutType 'hibernate-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$hibernateDc) {
            Write-Host "[SKIP] Hibernate After (Battery): already ${hibernateDc} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE hibernate-timeout-dc $hibernateDc | Out-Null
                Write-Host "[OK] Hibernate After (Battery): ${hibernateDc} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set hibernate (Battery)" -ForegroundColor Yellow
            }
        }
    }
}

# ========================================
# Button/Lid Action Settings Function
# ========================================
function Set-ButtonActions {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile
    )
    
    $activePlanGuid = Get-ActivePowerPlanGuid
    if (-not $activePlanGuid) {
        Write-Host "Warning: Skipping button settings because power plan GUID could not be retrieved" -ForegroundColor Yellow
        return
    }
    
    $buttonSubGroup = $script:SubGroupGuids['PowerButtons']
    
    # Power Button - AC
    $powerBtnAc = ConvertTo-SettingValue $Profile.PowerButton_AC
    if ($null -ne $powerBtnAc -and $script:ActionValues.ContainsKey($powerBtnAc)) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $buttonSubGroup `
            -SettingGuid $script:SettingGuids['PowerButton'] -Value $script:ActionValues[$powerBtnAc] `
            -PowerSource 'AC' -Description "Power Button (AC): $powerBtnAc"
    }
    
    # Power Button - Battery
    $powerBtnDc = ConvertTo-SettingValue $Profile.PowerButton_Battery
    if ($null -ne $powerBtnDc -and $script:ActionValues.ContainsKey($powerBtnDc)) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $buttonSubGroup `
            -SettingGuid $script:SettingGuids['PowerButton'] -Value $script:ActionValues[$powerBtnDc] `
            -PowerSource 'DC' -Description "Power Button (Battery): $powerBtnDc"
    }
    
    # Sleep Button - AC
    $sleepBtnAc = ConvertTo-SettingValue $Profile.SleepButton_AC
    if ($null -ne $sleepBtnAc -and $script:ActionValues.ContainsKey($sleepBtnAc)) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $buttonSubGroup `
            -SettingGuid $script:SettingGuids['SleepButton'] -Value $script:ActionValues[$sleepBtnAc] `
            -PowerSource 'AC' -Description "Sleep Button (AC): $sleepBtnAc"
    }
    
    # Sleep Button - Battery
    $sleepBtnDc = ConvertTo-SettingValue $Profile.SleepButton_Battery
    if ($null -ne $sleepBtnDc -and $script:ActionValues.ContainsKey($sleepBtnDc)) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $buttonSubGroup `
            -SettingGuid $script:SettingGuids['SleepButton'] -Value $script:ActionValues[$sleepBtnDc] `
            -PowerSource 'DC' -Description "Sleep Button (Battery): $sleepBtnDc"
    }
    
    # Lid Close - AC
    $lidAc = ConvertTo-SettingValue $Profile.LidClose_AC
    if ($null -ne $lidAc -and $script:ActionValues.ContainsKey($lidAc)) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $buttonSubGroup `
            -SettingGuid $script:SettingGuids['LidClose'] -Value $script:ActionValues[$lidAc] `
            -PowerSource 'AC' -Description "Lid Close (AC): $lidAc"
    }
    
    # Lid Close - Battery
    $lidDc = ConvertTo-SettingValue $Profile.LidClose_Battery
    if ($null -ne $lidDc -and $script:ActionValues.ContainsKey($lidDc)) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $buttonSubGroup `
            -SettingGuid $script:SettingGuids['LidClose'] -Value $script:ActionValues[$lidDc] `
            -PowerSource 'DC' -Description "Lid Close (Battery): $lidDc"
    }
}

# ========================================
# Hard Disk Settings Function
# ========================================
function Set-HardDiskSettings {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile,
        [string]$PlanGuid
    )

    # HDD - AC Power
    $hddAc = ConvertTo-SettingValue $Profile.HardDisk_TurnOff_AC
    if ($null -ne $hddAc) {
        $current = Get-TimeoutValue -TimeoutType 'disk-ac' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$hddAc) {
            Write-Host "[SKIP] HDD Turn Off (AC): already ${hddAc} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE disk-timeout-ac $hddAc | Out-Null
                Write-Host "[OK] HDD Turn Off (AC): ${hddAc} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set HDD (AC)" -ForegroundColor Yellow
            }
        }
    }

    # HDD - Battery
    $hddDc = ConvertTo-SettingValue $Profile.HardDisk_TurnOff_Battery
    if ($null -ne $hddDc) {
        $current = Get-TimeoutValue -TimeoutType 'disk-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$hddDc) {
            Write-Host "[SKIP] HDD Turn Off (Battery): already ${hddDc} min" -ForegroundColor Gray
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE disk-timeout-dc $hddDc | Out-Null
                Write-Host "[OK] HDD Turn Off (Battery): ${hddDc} min" -ForegroundColor Green
                $script:ChangeCount++
            }
            catch {
                Write-Host "[WARN] Failed to set HDD (Battery)" -ForegroundColor Yellow
            }
        }
    }
}

# ========================================
# Processor Settings Function
# ========================================
function Set-ProcessorSettings {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Profile
    )
    
    $activePlanGuid = Get-ActivePowerPlanGuid
    if (-not $activePlanGuid) {
        Write-Host "Warning: Skipping processor settings because power plan GUID could not be retrieved" -ForegroundColor Yellow
        return
    }
    
    $processorSubGroup = $script:SubGroupGuids['Processor']
    
    # Min Processor State - AC
    $minStateAc = ConvertTo-SettingValue $Profile.Processor_MinState_AC
    if ($null -ne $minStateAc) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $processorSubGroup `
            -SettingGuid $script:SettingGuids['ProcessorMinState'] -Value $minStateAc `
            -PowerSource 'AC' -Description "Min Processor State (AC): ${minStateAc}%"
    }
    
    # Min Processor State - Battery
    $minStateDc = ConvertTo-SettingValue $Profile.Processor_MinState_Battery
    if ($null -ne $minStateDc) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $processorSubGroup `
            -SettingGuid $script:SettingGuids['ProcessorMinState'] -Value $minStateDc `
            -PowerSource 'DC' -Description "Min Processor State (Battery): ${minStateDc}%"
    }
    
    # Max Processor State - AC
    $maxStateAc = ConvertTo-SettingValue $Profile.Processor_MaxState_AC
    if ($null -ne $maxStateAc) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $processorSubGroup `
            -SettingGuid $script:SettingGuids['ProcessorMaxState'] -Value $maxStateAc `
            -PowerSource 'AC' -Description "Max Processor State (AC): ${maxStateAc}%"
    }
    
    # Max Processor State - Battery
    $maxStateDc = ConvertTo-SettingValue $Profile.Processor_MaxState_Battery
    if ($null -ne $maxStateDc) {
        Set-PowerConfigValue -PlanGuid $activePlanGuid -SubGroupGuid $processorSubGroup `
            -SettingGuid $script:SettingGuids['ProcessorMaxState'] -Value $maxStateDc `
            -PowerSource 'DC' -Description "Max Processor State (Battery): ${maxStateDc}%"
    }
}

# ========================================
# powercfg Value Setting Helper Function
# ========================================
function Set-PowerConfigValue {
    param(
        [string]$PlanGuid,
        [string]$SubGroupGuid,
        [string]$SettingGuid,
        [int]$Value,
        [ValidateSet('AC', 'DC')]
        [string]$PowerSource,
        [string]$Description
    )

    # Idempotency check
    $current = Get-PowerConfigValue -PlanGuid $PlanGuid -SubGroupGuid $SubGroupGuid `
                -SettingGuid $SettingGuid -PowerSource $PowerSource
    if ($null -ne $current -and $current -eq $Value) {
        Write-Host "[SKIP] $Description (already set)" -ForegroundColor Gray
        $script:SkipCount++
        return
    }

    try {
        if ($PowerSource -eq 'AC') {
            $result = & powercfg /SETACVALUEINDEX $PlanGuid $SubGroupGuid $SettingGuid $Value 2>&1
        }
        else {
            $result = & powercfg /SETDCVALUEINDEX $PlanGuid $SubGroupGuid $SettingGuid $Value 2>&1
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] $Description" -ForegroundColor Green
            $script:ChangeCount++
        }
        else {
            Write-Host "[FAIL] $Description" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[ERROR] $Description" -ForegroundColor Yellow
    }
}

# ========================================
# Main Process
# ========================================
function Main {
    try {
        # Initialization
        Initialize-Script

        # Import CSV
        $profiles = Import-PowerSettingsCsv

        # Auto-selection: check for Enabled=1 entry
        $selectedProfile = $null
        $activeConfig = $profiles | Where-Object { $_.Enabled -eq '1' } | Select-Object -First 1

        if ($activeConfig) {
            $selectedProfile = $activeConfig
            Write-Host "[INFO] Auto-selected profile from CSV: " -NoNewline -ForegroundColor Cyan
            Write-Host "'$($selectedProfile.ProfileName)'" -ForegroundColor Yellow
            Write-Host ""
        } else {
            # Manual Selection (Fallback)
            $selectedProfile = Show-ProfileMenu -Profiles $profiles
        }

        if ($null -eq $selectedProfile) {
            Write-Host "Exiting process`n" -ForegroundColor Gray
            return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
        }

        # Confirmation
        if (-not (Confirm-ApplySettings -Profile $selectedProfile)) {
            Write-Host "Canceled`n" -ForegroundColor Gray
            return (New-ModuleResult -Status "Cancelled" -Message "User canceled")
        }
        
        Write-Host "`nApplying settings...`n" -ForegroundColor Cyan

        # Reset counters
        $script:SkipCount = 0
        $script:ChangeCount = 0

        # Change Power Plan
        $planValue = ConvertTo-SettingValue $selectedProfile.PowerPlan
        if ($null -ne $planValue) {
            Set-PowerPlan -PlanName $planValue
        }

        # Get active plan GUID (after potential plan change)
        $activePlanGuid = Get-ActivePowerPlanGuid

        # Apply Various Settings
        Set-DisplaySettings -Profile $selectedProfile -PlanGuid $activePlanGuid
        Set-SleepSettings -Profile $selectedProfile -PlanGuid $activePlanGuid
        Set-HardDiskSettings -Profile $selectedProfile -PlanGuid $activePlanGuid
        Set-ButtonActions -Profile $selectedProfile
        Set-ProcessorSettings -Profile $selectedProfile

        # Activate Settings
        Write-Host "`nActivating settings..." -ForegroundColor Gray
        & powercfg /SETACTIVE (Get-ActivePowerPlanGuid) | Out-Null

        Write-Host "`n========================================" -ForegroundColor Green
        Write-Host "  Configuration Results" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  Changed: $($script:ChangeCount), Skipped: $($script:SkipCount)" -ForegroundColor Green
        Write-Host "========================================`n" -ForegroundColor Green

        if ($script:ChangeCount -eq 0 -and $script:SkipCount -gt 0) {
            return (New-ModuleResult -Status "Skipped" -Message "All $($script:SkipCount) settings already configured")
        }
        return (New-ModuleResult -Status "Success" -Message "Changed: $($script:ChangeCount), Skip: $($script:SkipCount)")
    }
    catch {
        Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Error occurred while applying settings`n" -ForegroundColor Red
        return (New-ModuleResult -Status "Error" -Message "Power settings failed: $($_.Exception.Message)")
    }
}

Main