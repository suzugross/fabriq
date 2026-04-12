# ========================================
# Volume Configuration Script
# ========================================
# [PURPOSE]
# Set master volume and mute state via Windows Core Audio API.
# Uses IAudioEndpointVolume COM interface for immediate effect.
# ========================================

Write-Host ""
Show-Separator
Write-Host "Volume Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Core Audio API COM Interop Definition
# ========================================
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class VolumeHandler
{
    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    private class MMDeviceEnumerator {}

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceEnumerator
    {
        int NotImpl_EnumAudioEndpoints();
        int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice device);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDevice
    {
        int Activate(ref Guid iid, int clsCtx, IntPtr activationParams,
                     [MarshalAs(UnmanagedType.IUnknown)] out object volume);
    }

    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IAudioEndpointVolume
    {
        int RegisterControlChangeNotify(IntPtr client);
        int UnregisterControlChangeNotify(IntPtr client);
        int GetChannelCount(out uint count);
        int SetMasterVolumeLevel(float level, ref Guid eventContext);
        int SetMasterVolumeLevelScalar(float level, ref Guid eventContext);
        int GetMasterVolumeLevel(out float level);
        int GetMasterVolumeLevelScalar(out float level);
        int SetChannelVolumeLevel(uint ch, float level, ref Guid ctx);
        int SetChannelVolumeLevelScalar(uint ch, float level, ref Guid ctx);
        int GetChannelVolumeLevel(uint ch, out float level);
        int GetChannelVolumeLevelScalar(uint ch, out float level);
        int SetMute([MarshalAs(UnmanagedType.Bool)] bool mute, ref Guid eventContext);
        int GetMute([MarshalAs(UnmanagedType.Bool)] out bool mute);
    }

    private static readonly Guid IID_IAudioEndpointVolume =
        new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");

    private static IAudioEndpointVolume GetEndpointVolume()
    {
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
        IMMDevice device;
        // eRender=0, eMultimedia=1
        enumerator.GetDefaultAudioEndpoint(0, 1, out device);
        object obj;
        Guid iid = IID_IAudioEndpointVolume;
        device.Activate(ref iid, 1, IntPtr.Zero, out obj);
        return (IAudioEndpointVolume)obj;
    }

    public static int GetVolume()
    {
        var vol = GetEndpointVolume();
        float level;
        vol.GetMasterVolumeLevelScalar(out level);
        return (int)Math.Round(level * 100);
    }

    public static void SetVolume(int percent)
    {
        var vol = GetEndpointVolume();
        float level = percent / 100.0f;
        Guid guid = Guid.Empty;
        vol.SetMasterVolumeLevelScalar(level, ref guid);
    }

    public static bool GetMute()
    {
        var vol = GetEndpointVolume();
        bool mute;
        vol.GetMute(out mute);
        return mute;
    }

    public static void SetMute(bool mute)
    {
        var vol = GetEndpointVolume();
        Guid guid = Guid.Empty;
        vol.SetMute(mute, ref guid);
    }
}
'@ -ErrorAction SilentlyContinue

# ========================================
# Step 1: CSV Load
# ========================================
$csvPath = Join-Path $PSScriptRoot "volume_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Volume")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load volume_list.csv")
}

$config = @($enabledItems)[0]

$targetVolume = 0
if (-not [int]::TryParse($config.Volume, [ref]$targetVolume) -or $targetVolume -lt 0 -or $targetVolume -gt 100) {
    Show-Error "Invalid Volume value: $($config.Volume) (must be 0-100)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Invalid Volume value: $($config.Volume)")
}

$targetMute = $null
if ($config.PSObject.Properties.Name -contains 'Mute' -and -not [string]::IsNullOrWhiteSpace($config.Mute)) {
    $muteValue = $config.Mute.Trim().ToLower()
    if ($muteValue -eq 'on') {
        $targetMute = $true
    }
    elseif ($muteValue -eq 'off') {
        $targetMute = $false
    }
    else {
        Show-Warning "Unknown Mute value: $($config.Mute) (ignoring)"
    }
}

# ========================================
# Step 2: Prerequisite Check
# ========================================
$currentVolume = -1
$currentMute = $false

try {
    $currentVolume = [VolumeHandler]::GetVolume()
    $currentMute = [VolumeHandler]::GetMute()
}
catch {
    Show-Error "No audio device found or Core Audio API unavailable: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Audio device not available")
}

# ========================================
# Step 3: Display Current & Target State
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Volume Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

$currentMuteText = if ($currentMute) { "ON (Muted)" } else { "OFF" }
Write-Host "  Current Volume: $currentVolume%" -ForegroundColor White
Write-Host "  Current Mute:   $currentMuteText" -ForegroundColor White
Write-Host ""

$targetMuteText = if ($null -eq $targetMute) { "(no change)" } elseif ($targetMute) { "ON (Muted)" } else { "OFF" }
Write-Host "  Target Volume:  $targetVolume%" -ForegroundColor Cyan
Write-Host "  Target Mute:    $targetMuteText" -ForegroundColor Cyan
Write-Host ""

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

# Idempotency check
$volumeMatch = ($currentVolume -eq $targetVolume)
$muteMatch = ($null -eq $targetMute) -or ($currentMute -eq $targetMute)

if ($volumeMatch -and $muteMatch) {
    Show-Skip "Already at target state (Volume: $targetVolume%, Mute: $currentMuteText)"
    Write-Host ""
    return (New-ModuleResult -Status "Skipped" -Message "Already at target state" -Verified $true)
}

# ========================================
# Step 4: Confirmation
# ========================================
$changes = @()
if (-not $volumeMatch) { $changes += "Volume: $currentVolume% -> $targetVolume%" }
if (-not $muteMatch)   { $changes += "Mute: $currentMuteText -> $targetMuteText" }

Show-Info "Changes: $($changes -join ', ')"
Write-Host ""

$cancelResult = Confirm-ModuleExecution -Message "Apply volume settings?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""

# ========================================
# Step 5: Apply Settings
# ========================================
$successCount = 0
$failCount = 0

if (-not $volumeMatch) {
    Show-Info "Setting volume to $targetVolume%..."
    try {
        [VolumeHandler]::SetVolume($targetVolume)
        Show-Success "Volume set to $targetVolume%"
        $successCount++
    }
    catch {
        Show-Error "Failed to set volume: $_"
        $failCount++
    }
}

if (-not $muteMatch) {
    $muteAction = if ($targetMute) { "Muting" } else { "Unmuting" }
    Show-Info "$muteAction audio..."
    try {
        [VolumeHandler]::SetMute($targetMute)
        Show-Success "$muteAction complete"
        $successCount++
    }
    catch {
        Show-Error "Failed to set mute: $_"
        $failCount++
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

try {
    $actualVolume = [VolumeHandler]::GetVolume()
    $actualMute = [VolumeHandler]::GetMute()

    if ($actualVolume -eq $targetVolume) {
        Write-Host "  [VERIFIED] Volume: $actualVolume%" -ForegroundColor Green
        $verifyPass++
    }
    else {
        Write-Host "  [VERIFY FAILED] Volume: expected $targetVolume%, got $actualVolume%" -ForegroundColor Red
        $verifyFail++
    }

    if ($null -ne $targetMute) {
        $actualMuteText = if ($actualMute) { "ON" } else { "OFF" }
        if ($actualMute -eq $targetMute) {
            Write-Host "  [VERIFIED] Mute: $actualMuteText" -ForegroundColor Green
            $verifyPass++
        }
        else {
            Write-Host "  [VERIFY FAILED] Mute: expected $targetMuteText, got $actualMuteText" -ForegroundColor Red
            $verifyFail++
        }
    }
}
catch {
    Show-Warning "Failed to verify: $_"
    $verifyFail++
}

Write-Host ""
$verified = ($verifyFail -eq 0)

# ========================================
# Step 6: Result Summary
# ========================================
return (New-BatchResult -Success $successCount -Skip 0 -Fail $failCount `
    -Title "Volume Configuration Results" -Verified $verified)
