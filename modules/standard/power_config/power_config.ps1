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

# Power Mode Overlay GUIDs (Windows performance power slider)
# These overlay the base power plan to adjust performance vs battery tradeoff.
# Reference: https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/customize-power-slider
$script:PowerModeGuids = @{
    'BEST_EFFICIENCY'   = '961cc777-2547-4f9d-8174-7d86181b8a7a'
    'BALANCED'          = '00000000-0000-0000-0000-000000000000'
    'BEST_PERFORMANCE'  = 'ded574b5-45a0-4f42-8737-46345c09c238'
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

# P/Invoke: powrprof.dll Power Mode APIs (Windows 11+)
# These APIs are the same code path used by Windows Settings GUI,
# allowing AC/DC overlay to be set independently without direct registry writes.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class PowerModeApi {
    [DllImport("powrprof.dll", EntryPoint = "PowerSetUserConfiguredACPowerMode")]
    public static extern uint SetACPowerMode(ref Guid PowerModeGuid);

    [DllImport("powrprof.dll", EntryPoint = "PowerSetUserConfiguredDCPowerMode")]
    public static extern uint SetDCPowerMode(ref Guid PowerModeGuid);

    [DllImport("powrprof.dll", EntryPoint = "PowerGetUserConfiguredACPowerMode")]
    public static extern uint GetACPowerMode(out Guid PowerModeGuid);

    [DllImport("powrprof.dll", EntryPoint = "PowerGetUserConfiguredDCPowerMode")]
    public static extern uint GetDCPowerMode(out Guid PowerModeGuid);
}
"@ -ErrorAction SilentlyContinue

# P/Invoke: powrprof.dll Power Write/Read APIs
# These APIs are the same code path used by Control Panel UI,
# bypassing OEM restrictions that affect powercfg.exe on some PCs (e.g., HP).
# Read APIs are locale-independent (unlike powercfg /QUERY text output).
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class PowerWriteApi {
    [DllImport("powrprof.dll")]
    public static extern uint PowerWriteACValueIndex(
        IntPtr RootPowerKey,
        ref Guid SchemeGuid,
        ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid,
        uint AcValueIndex);

    [DllImport("powrprof.dll")]
    public static extern uint PowerWriteDCValueIndex(
        IntPtr RootPowerKey,
        ref Guid SchemeGuid,
        ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid,
        uint DcValueIndex);

    [DllImport("powrprof.dll")]
    public static extern uint PowerReadACValueIndex(
        IntPtr RootPowerKey,
        ref Guid SchemeGuid,
        ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid,
        out uint AcValueIndex);

    [DllImport("powrprof.dll")]
    public static extern uint PowerReadDCValueIndex(
        IntPtr RootPowerKey,
        ref Guid SchemeGuid,
        ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid,
        out uint DcValueIndex);

    [DllImport("powrprof.dll")]
    public static extern uint PowerSetActiveScheme(
        IntPtr UserRootPowerKey,
        ref Guid SchemeGuid);
}
"@ -ErrorAction SilentlyContinue

# Idempotency counters
$script:SkipCount = 0
$script:ChangeCount = 0
# Deterministic failures only (explicit powercfg non-zero exit / CSV config
# errors). Hardware-dependent Win32 hr!=0 paths (power mode overlay, button
# and lid writes) stay warnings by design: they fail benignly on machines
# without the capability, and Step 5.5 verification (Verified) is the
# backstop for them. See TM t-0017.
$script:FailCount = 0

# ========================================
# Initialization Function
# ========================================
function Initialize-Script {
    Write-Host "Starting Power Option Configuration Script`n" -ForegroundColor Cyan
}

# ========================================
# CSV Import Function
# ========================================
function Import-PowerSettingsCsv {
    $csvData = Import-ModuleCsv -Path $script:CsvPath
    if ($null -eq $csvData -or $csvData.Count -eq 0) {
        throw "Failed to load power_list.csv"
    }

    Write-Host "Loaded $($csvData.Count) profiles`n" -ForegroundColor Green
    return $csvData
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
# Uses PowerReadACValueIndex/PowerReadDCValueIndex Win32 APIs directly instead of
# parsing powercfg /QUERY text output, which is localized and breaks on non-English
# Windows (e.g., Japanese). This API is the canonical read counterpart of
# PowerWriteACValueIndex/PowerWriteDCValueIndex used for SET, ensuring SET/GET path
# consistency and full locale independence.
function Get-PowerConfigValue {
    param(
        [string]$PlanGuid,
        [string]$SubGroupGuid,
        [string]$SettingGuid,
        [ValidateSet('AC', 'DC')]
        [string]$PowerSource
    )

    try {
        $scheme = [Guid]$PlanGuid
        $sub    = [Guid]$SubGroupGuid
        $sett   = [Guid]$SettingGuid
        $val    = [uint32]0

        if ($PowerSource -eq 'AC') {
            $hr = [PowerWriteApi]::PowerReadACValueIndex(
                [IntPtr]::Zero, [ref]$scheme, [ref]$sub, [ref]$sett, [ref]$val)
        }
        else {
            $hr = [PowerWriteApi]::PowerReadDCValueIndex(
                [IntPtr]::Zero, [ref]$scheme, [ref]$sub, [ref]$sett, [ref]$val)
        }

        if ($hr -eq 0) {
            return [int64]$val
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
    
    Show-Separator
    Write-Host "  Select Power Option Profile" -ForegroundColor Cyan
    Show-Separator
    Write-Host ""
    
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
    
    Show-Separator
    Write-Host "  Settings to Apply" -ForegroundColor Cyan
    Show-Separator
    Write-Host "Profile Name: " -NoNewline
    Write-Host $Profile.ProfileName -ForegroundColor Yellow
    Write-Host "Description: " -NoNewline
    Write-Host $Profile.Description -ForegroundColor Gray
    Write-Host "Power Plan: " -NoNewline
    Write-Host $Profile.PowerPlan -ForegroundColor Green
    $modeValue = ConvertTo-SettingValue $Profile.PowerMode
    if ($null -ne $modeValue) {
        Write-Host "Power Mode: " -NoNewline
        Write-Host $modeValue -ForegroundColor Green
    }
    Show-Separator
    Write-Host ""

    return (Confirm-Execution -Message "Apply these settings?")
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
            # CSV config error (deterministic) - counts as Fail
            Show-Warning "Unknown power plan: $PlanName"
            $script:FailCount++
            return $false
        }

        # Idempotency check
        $currentGuid = Get-ActivePowerPlanGuid
        if ($currentGuid -eq $planGuid) {
            Show-Skip "Power plan already '$PlanName'"
            $script:SkipCount++
            return $true
        }

        Write-Host "Changing power plan to '$PlanName'..." -ForegroundColor Gray

        $result = & powercfg /S $planGuid 2>&1

        if ($LASTEXITCODE -eq 0) {
            Show-Success "Changed power plan to '$PlanName'"
            $script:ChangeCount++
            return $true
        }
        else {
            Show-Error "Failed to change power plan: $result"
            $script:FailCount++
            return $false
        }
    }
    catch {
        Show-Error "Failed to set power plan - $($_.Exception.Message)"
        $script:FailCount++
        return $false
    }
}

# ========================================
# Power Mode Overlay Setting Function
# ========================================
function Set-PowerMode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModeName
    )

    try {
        $modeGuidStr = $script:PowerModeGuids[$ModeName]

        if (-not $modeGuidStr) {
            # CSV config error (deterministic) - counts as Fail
            Show-Warning "Unknown power mode: $ModeName"
            $script:FailCount++
            return $false
        }

        $targetGuid = [Guid]::new($modeGuidStr)

        # Apply overlay for AC and DC via powrprof.dll API
        foreach ($source in @('AC', 'DC')) {
            $description = "Power mode ($source): $ModeName"

            # Idempotency check via Getter API. Trust the skip only when the
            # read itself succeeded: a failed read leaves [Guid]::Empty,
            # which equals the BALANCED overlay GUID (all zeros) and would
            # false-skip a BALANCED target. On read failure fall through to
            # apply (fail-closed toward the apply side).
            $currentGuid = [Guid]::Empty
            $readHr = if ($source -eq 'AC') {
                [PowerModeApi]::GetACPowerMode([ref]$currentGuid)
            }
            else {
                [PowerModeApi]::GetDCPowerMode([ref]$currentGuid)
            }

            if ($readHr -eq 0 -and $currentGuid -eq $targetGuid) {
                Show-Skip "$description (already set)"
                $script:SkipCount++
                continue
            }

            # Apply via Setter API
            $setGuid = $targetGuid
            if ($source -eq 'AC') {
                $hr = [PowerModeApi]::SetACPowerMode([ref]$setGuid)
            }
            else {
                $hr = [PowerModeApi]::SetDCPowerMode([ref]$setGuid)
            }

            if ($hr -eq 0) {
                Show-Success "$description"
                $script:ChangeCount++
            }
            else {
                # Warning only (no FailCount): power mode overlays are
                # unsupported on non-Modern-Standby machines, where this
                # hr is a benign hardware limitation. Step 5.5 (Verified)
                # is the backstop.
                Show-Warning "Failed to set $description (error code: $hr)"
            }
        }

        return $true
    }
    catch {
        Show-Error "Failed to set power mode - $($_.Exception.Message)"
        $script:FailCount++
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
        Show-Warning "Failed to get current power plan"
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
            Show-Skip "Display Turn Off (AC): already ${acValue} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE monitor-timeout-ac $acValue | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "Display Turn Off (AC): ${acValue} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "Display Turn Off (AC) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set display (AC)"
                $script:FailCount++
            }
        }
    }

    # On Battery
    $batteryValue = ConvertTo-SettingValue $Profile.Display_TurnOff_Battery
    if ($null -ne $batteryValue) {
        $current = Get-TimeoutValue -TimeoutType 'monitor-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$batteryValue) {
            Show-Skip "Display Turn Off (Battery): already ${batteryValue} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE monitor-timeout-dc $batteryValue | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "Display Turn Off (Battery): ${batteryValue} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "Display Turn Off (Battery) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set display (Battery)"
                $script:FailCount++
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
            Show-Skip "Sleep After (AC): already ${sleepAc} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE standby-timeout-ac $sleepAc | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "Sleep After (AC): ${sleepAc} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "Sleep After (AC) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set sleep (AC)"
                $script:FailCount++
            }
        }
    }

    # Sleep - Battery
    $sleepDc = ConvertTo-SettingValue $Profile.Sleep_After_Battery
    if ($null -ne $sleepDc) {
        $current = Get-TimeoutValue -TimeoutType 'standby-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$sleepDc) {
            Show-Skip "Sleep After (Battery): already ${sleepDc} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE standby-timeout-dc $sleepDc | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "Sleep After (Battery): ${sleepDc} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "Sleep After (Battery) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set sleep (Battery)"
                $script:FailCount++
            }
        }
    }

    # Hibernate - AC Power
    $hibernateAc = ConvertTo-SettingValue $Profile.Hibernate_After_AC
    if ($null -ne $hibernateAc) {
        $current = Get-TimeoutValue -TimeoutType 'hibernate-ac' -PlanGuid $PlanGuid
        if ($null -ne $current -and [math]::Abs($current - [int]$hibernateAc) -le 1) {
            Show-Skip "Hibernate After (AC): already ${hibernateAc} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE hibernate-timeout-ac $hibernateAc | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "Hibernate After (AC): ${hibernateAc} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "Hibernate After (AC) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set hibernate (AC)"
                $script:FailCount++
            }
        }
    }

    # Hibernate - Battery
    $hibernateDc = ConvertTo-SettingValue $Profile.Hibernate_After_Battery
    if ($null -ne $hibernateDc) {
        $current = Get-TimeoutValue -TimeoutType 'hibernate-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and [math]::Abs($current - [int]$hibernateDc) -le 1) {
            Show-Skip "Hibernate After (Battery): already ${hibernateDc} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE hibernate-timeout-dc $hibernateDc | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "Hibernate After (Battery): ${hibernateDc} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "Hibernate After (Battery) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set hibernate (Battery)"
                $script:FailCount++
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
        Show-Warning "Skipping button settings because power plan GUID could not be retrieved"
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

    # Apply changes immediately via Win32 API (equivalent to powercfg /SETACTIVE)
    $schemeGuid = [Guid]$activePlanGuid
    $hr = [PowerWriteApi]::PowerSetActiveScheme([IntPtr]::Zero, [ref]$schemeGuid)
    if ($hr -ne 0) {
        Show-Warning "PowerSetActiveScheme failed: 0x$($hr.ToString('X8'))"
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
            Show-Skip "HDD Turn Off (AC): already ${hddAc} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE disk-timeout-ac $hddAc | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "HDD Turn Off (AC): ${hddAc} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "HDD Turn Off (AC) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set HDD (AC)"
                $script:FailCount++
            }
        }
    }

    # HDD - Battery
    $hddDc = ConvertTo-SettingValue $Profile.HardDisk_TurnOff_Battery
    if ($null -ne $hddDc) {
        $current = Get-TimeoutValue -TimeoutType 'disk-dc' -PlanGuid $PlanGuid
        if ($null -ne $current -and $current -eq [int]$hddDc) {
            Show-Skip "HDD Turn Off (Battery): already ${hddDc} min"
            $script:SkipCount++
        }
        else {
            try {
                & powercfg /CHANGE disk-timeout-dc $hddDc | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Show-Success "HDD Turn Off (Battery): ${hddDc} min"
                    $script:ChangeCount++
                }
                else {
                    Show-Error "HDD Turn Off (Battery) failed (ExitCode=$LASTEXITCODE)"
                    $script:FailCount++
                }
            }
            catch {
                Show-Warning "Failed to set HDD (Battery)"
                $script:FailCount++
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
        Show-Warning "Skipping processor settings because power plan GUID could not be retrieved"
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
        Show-Skip "$Description (already set)"
        $script:SkipCount++
        return
    }

    try {
        $isButtonSetting = ($SubGroupGuid -eq $script:SubGroupGuids['PowerButtons'])

        if ($isButtonSetting) {
            # Use Win32 API for button/lid settings (same code path as Control Panel UI)
            # This bypasses OEM restrictions that cause powercfg.exe to silently fail
            $schemeGuid   = [Guid]$PlanGuid
            $subGuid      = [Guid]$SubGroupGuid
            $settGuid     = [Guid]$SettingGuid

            if ($PowerSource -eq 'AC') {
                $hr = [PowerWriteApi]::PowerWriteACValueIndex(
                    [IntPtr]::Zero, [ref]$schemeGuid, [ref]$subGuid,
                    [ref]$settGuid, [uint32]$Value)
            }
            else {
                $hr = [PowerWriteApi]::PowerWriteDCValueIndex(
                    [IntPtr]::Zero, [ref]$schemeGuid, [ref]$subGuid,
                    [ref]$settGuid, [uint32]$Value)
            }

            if ($hr -eq 0) {
                Show-Success "$Description"
                $script:ChangeCount++
            }
            else {
                # Warning only (no FailCount): button/lid writes can fail
                # benignly on hardware without the capability (e.g. lid
                # settings on desktops). Step 5.5 (Verified) is the backstop.
                Show-Warning "$Description (Win32 API failed: 0x$($hr.ToString('X8')))"
            }
        }
        else {
            # Use powercfg for other settings (Display, Sleep, HDD, Processor)
            if ($PowerSource -eq 'AC') {
                $result = & powercfg /SETACVALUEINDEX $PlanGuid $SubGroupGuid $SettingGuid $Value 2>&1
            }
            else {
                $result = & powercfg /SETDCVALUEINDEX $PlanGuid $SubGroupGuid $SettingGuid $Value 2>&1
            }

            if ($LASTEXITCODE -eq 0) {
                Show-Success "$Description"
                $script:ChangeCount++
            }
            else {
                # Explicit powercfg non-zero exit = deterministic failure
                Show-Error "$Description (failed, ExitCode=$LASTEXITCODE)"
                $script:FailCount++
            }
        }
    }
    catch {
        Show-Error "$Description"
        $script:FailCount++
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
        $script:FailCount = 0

        # Change Power Plan
        $planValue = ConvertTo-SettingValue $selectedProfile.PowerPlan
        if ($null -ne $planValue) {
            Set-PowerPlan -PlanName $planValue
        }

        # Change Power Mode Overlay
        $modeValue = ConvertTo-SettingValue $selectedProfile.PowerMode
        if ($null -ne $modeValue) {
            Set-PowerMode -ModeName $modeValue
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
        if ($LASTEXITCODE -ne 0) {
            Show-Error "powercfg /SETACTIVE failed (ExitCode=$LASTEXITCODE)"
            $script:FailCount++
        }

        # ========================================
        # Step 5.5: Post-Apply Verification
        # ========================================
        Write-Host ""
        Show-Info "Verifying applied settings..."
        Write-Host ""

        $verifyPass = 0
        $verifyFail = 0

        # Re-read active plan GUID for verification
        $verifyPlanGuid = Get-ActivePowerPlanGuid

        # --- Power Plan ---
        $planTarget = ConvertTo-SettingValue $selectedProfile.PowerPlan
        if ($null -ne $planTarget -and $script:PowerPlanGuids.ContainsKey($planTarget)) {
            $expectedGuid = $script:PowerPlanGuids[$planTarget]
            if ($verifyPlanGuid -eq $expectedGuid) {
                Write-Host "  [VERIFIED] Power Plan: $planTarget" -ForegroundColor Green
                $verifyPass++
            } else {
                Write-Host "  [VERIFY FAILED] Power Plan: expected $planTarget ($expectedGuid), actual $verifyPlanGuid" -ForegroundColor Red
                $verifyFail++
            }
        }

        # --- Power Mode Overlay ---
        $modeTarget = ConvertTo-SettingValue $selectedProfile.PowerMode
        if ($null -ne $modeTarget -and $script:PowerModeGuids.ContainsKey($modeTarget)) {
            $expectedModeGuid = [Guid]::new($script:PowerModeGuids[$modeTarget])

            foreach ($src in @('AC', 'DC')) {
                $currentModeGuid = [Guid]::Empty
                $hr = if ($src -eq 'AC') {
                    [PowerModeApi]::GetACPowerMode([ref]$currentModeGuid)
                } else {
                    [PowerModeApi]::GetDCPowerMode([ref]$currentModeGuid)
                }

                # A failed read leaves the sentinel [Guid]::Empty, which
                # collides with the BALANCED overlay GUID (all zeros): a
                # discarded HRESULT turned an unreadable overlay into a
                # false PASS for BALANCED targets (and a false FAIL for
                # everything else). Overlays are hardware-dependent (the
                # apply side warns without failing), so an unreadable
                # overlay is "unverifiable", not pass or fail.
                if ($hr -ne 0) {
                    Write-Host "  [VERIFY SKIPPED] Power Mode ($src): overlay not readable on this hardware (hr=0x$('{0:X8}' -f $hr))" -ForegroundColor Yellow
                    continue
                }

                if ($currentModeGuid -eq $expectedModeGuid) {
                    Write-Host "  [VERIFIED] Power Mode ($src): $modeTarget" -ForegroundColor Green
                    $verifyPass++
                } else {
                    Write-Host "  [VERIFY FAILED] Power Mode ($src): expected $modeTarget, actual $currentModeGuid" -ForegroundColor Red
                    $verifyFail++
                }
            }
        }

        # --- Timeout Settings ---
        $timeoutChecks = @(
            @{ Property = 'Display_TurnOff_AC';       Type = 'monitor-ac';   Label = 'Display Off (AC)' }
            @{ Property = 'Display_TurnOff_Battery';   Type = 'monitor-dc';   Label = 'Display Off (DC)' }
            @{ Property = 'Sleep_After_AC';            Type = 'standby-ac';   Label = 'Sleep (AC)' }
            @{ Property = 'Sleep_After_Battery';       Type = 'standby-dc';   Label = 'Sleep (DC)' }
            @{ Property = 'Hibernate_After_AC';        Type = 'hibernate-ac'; Label = 'Hibernate (AC)' }
            @{ Property = 'Hibernate_After_Battery';   Type = 'hibernate-dc'; Label = 'Hibernate (DC)' }
            @{ Property = 'HardDisk_TurnOff_AC';       Type = 'disk-ac';      Label = 'HDD Off (AC)' }
            @{ Property = 'HardDisk_TurnOff_Battery';  Type = 'disk-dc';      Label = 'HDD Off (DC)' }
        )

        foreach ($tc in $timeoutChecks) {
            $targetVal = ConvertTo-SettingValue $selectedProfile.($tc.Property)
            if ($null -ne $targetVal) {
                $currentMin = Get-TimeoutValue -TimeoutType $tc.Type -PlanGuid $verifyPlanGuid
                if ($null -ne $currentMin -and $currentMin -eq [int]$targetVal) {
                    # Exact match
                    Write-Host "  [VERIFIED] $($tc.Label): $targetVal min" -ForegroundColor Green
                    $verifyPass++
                } elseif ($null -ne $currentMin -and [math]::Abs($currentMin - [int]$targetVal) -le 1) {
                    # Within ±1 min tolerance (Windows internal adjustment)
                    Write-Host "  [VERIFIED] $($tc.Label): $targetVal min (actual: $currentMin min, within tolerance)" -ForegroundColor Green
                    $verifyPass++
                } else {
                    $actual = if ($null -ne $currentMin) { "$currentMin min" } else { "Unknown" }
                    Write-Host "  [VERIFY FAILED] $($tc.Label): expected $targetVal min, actual $actual" -ForegroundColor Red
                    $verifyFail++
                }
            }
        }

        # --- Button/Lid Actions ---
        $buttonSubGroup = $script:SubGroupGuids['PowerButtons']
        $buttonChecks = @(
            @{ Property = 'PowerButton_AC';   Guid = $script:SettingGuids['PowerButton']; Src = 'AC'; Label = 'Power Button (AC)' }
            @{ Property = 'PowerButton_Battery'; Guid = $script:SettingGuids['PowerButton']; Src = 'DC'; Label = 'Power Button (DC)' }
            @{ Property = 'SleepButton_AC';   Guid = $script:SettingGuids['SleepButton']; Src = 'AC'; Label = 'Sleep Button (AC)' }
            @{ Property = 'SleepButton_Battery'; Guid = $script:SettingGuids['SleepButton']; Src = 'DC'; Label = 'Sleep Button (DC)' }
            @{ Property = 'LidClose_AC';      Guid = $script:SettingGuids['LidClose'];    Src = 'AC'; Label = 'Lid Close (AC)' }
            @{ Property = 'LidClose_Battery'; Guid = $script:SettingGuids['LidClose'];    Src = 'DC'; Label = 'Lid Close (DC)' }
        )

        foreach ($bc in $buttonChecks) {
            $targetName = ConvertTo-SettingValue $selectedProfile.($bc.Property)
            if ($null -ne $targetName -and $script:ActionValues.ContainsKey($targetName)) {
                $expectedVal = $script:ActionValues[$targetName]
                $currentVal = Get-PowerConfigValue -PlanGuid $verifyPlanGuid -SubGroupGuid $buttonSubGroup `
                    -SettingGuid $bc.Guid -PowerSource $bc.Src

                if ($null -ne $currentVal -and $currentVal -eq $expectedVal) {
                    Write-Host "  [VERIFIED] $($bc.Label): $targetName" -ForegroundColor Green
                    $verifyPass++
                } else {
                    $actual = if ($null -ne $currentVal) { $currentVal } else { "Unknown" }
                    Write-Host "  [VERIFY FAILED] $($bc.Label): expected $targetName ($expectedVal), actual $actual" -ForegroundColor Red
                    $verifyFail++
                }
            }
        }

        # --- Processor States ---
        $processorSubGroup = $script:SubGroupGuids['Processor']
        $procChecks = @(
            @{ Property = 'Processor_MinState_AC';      Guid = $script:SettingGuids['ProcessorMinState']; Src = 'AC'; Label = 'Min CPU (AC)' }
            @{ Property = 'Processor_MinState_Battery';  Guid = $script:SettingGuids['ProcessorMinState']; Src = 'DC'; Label = 'Min CPU (DC)' }
            @{ Property = 'Processor_MaxState_AC';      Guid = $script:SettingGuids['ProcessorMaxState']; Src = 'AC'; Label = 'Max CPU (AC)' }
            @{ Property = 'Processor_MaxState_Battery';  Guid = $script:SettingGuids['ProcessorMaxState']; Src = 'DC'; Label = 'Max CPU (DC)' }
        )

        foreach ($pc in $procChecks) {
            $targetVal = ConvertTo-SettingValue $selectedProfile.($pc.Property)
            if ($null -ne $targetVal) {
                $currentVal = Get-PowerConfigValue -PlanGuid $verifyPlanGuid -SubGroupGuid $processorSubGroup `
                    -SettingGuid $pc.Guid -PowerSource $pc.Src

                if ($null -ne $currentVal -and $currentVal -eq [int]$targetVal) {
                    Write-Host "  [VERIFIED] $($pc.Label): $targetVal%" -ForegroundColor Green
                    $verifyPass++
                } else {
                    $actual = if ($null -ne $currentVal) { "$currentVal%" } else { "Unknown" }
                    Write-Host "  [VERIFY FAILED] $($pc.Label): expected $targetVal%, actual $actual" -ForegroundColor Red
                    $verifyFail++
                }
            }
        }

        Write-Host ""
        $verified = if ($verifyPass -eq 0 -and $verifyFail -eq 0) { $null } else { $verifyFail -eq 0 }

        $bannerColor = if ($script:FailCount -gt 0) { "Yellow" } else { "Green" }
        Write-Host "========================================" -ForegroundColor $bannerColor
        Write-Host "  Configuration Results" -ForegroundColor $bannerColor
        Write-Host "========================================" -ForegroundColor $bannerColor
        Write-Host "  Changed: $($script:ChangeCount), Skipped: $($script:SkipCount)" -ForegroundColor $bannerColor
        if ($script:FailCount -gt 0) {
            Write-Host "  Failed:  $($script:FailCount)" -ForegroundColor Red
        }
        Write-Host "========================================`n" -ForegroundColor $bannerColor

        # FailCount holds deterministic failures only (explicit powercfg
        # non-zero exits / CSV config errors); hardware-dependent Win32
        # warnings are excluded by design and covered by Verified.
        if ($script:FailCount -gt 0 -and $script:ChangeCount -eq 0) {
            return (New-ModuleResult -Status "Error" -Message "Changed: $($script:ChangeCount), Skip: $($script:SkipCount), Fail: $($script:FailCount)" -Verified $verified)
        }
        if ($script:FailCount -gt 0) {
            return (New-ModuleResult -Status "Partial" -Message "Changed: $($script:ChangeCount), Skip: $($script:SkipCount), Fail: $($script:FailCount)" -Verified $verified)
        }
        if ($script:ChangeCount -eq 0 -and $script:SkipCount -gt 0) {
            return (New-ModuleResult -Status "Skipped" -Message "All $($script:SkipCount) settings already configured" -Verified $verified)
        }
        return (New-ModuleResult -Status "Success" -Message "Changed: $($script:ChangeCount), Skip: $($script:SkipCount)" -Verified $verified)
    }
    catch {
        Write-Host ""
        Show-Error "$($_.Exception.Message)"
        Show-Error "Error occurred while applying settings"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Power settings failed: $($_.Exception.Message)")
    }
}

Main