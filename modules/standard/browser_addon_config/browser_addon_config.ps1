# ========================================
# Browser Addon Configuration Script
# ========================================
# Chrome/Edge の拡張機能をグループポリシー（レジストリ）経由で
# 強制インストールする。ExtensionInstallForcelist に登録。
#
# [NOTES]
# - 管理者権限が必要
# - ブラウザの再起動または gpupdate /force が必要な場合あり
# ========================================

Write-Host ""
Show-Separator
Write-Host "Browser Addon Configuration" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# Local Helper Functions
# ========================================

function Resolve-ExtensionId {
    param([string]$InputValue)
    $InputValue = $InputValue.Trim()
    # Raw extension ID (32 chars, a-p only)
    if ($InputValue -match '^[a-p]{32}$') { return $InputValue }
    # Chrome Web Store URL: /detail/[optional-name-slug/]<id>
    if ($InputValue -match '/detail/(?:[^/]+/)?([a-p]{32})') { return $Matches[1] }
    # Fallback: any 32-char a-p sequence in the string
    if ($InputValue -match '([a-p]{32})') { return $Matches[1] }
    return $null
}

function Test-ExtensionInForcelist {
    param(
        [string]$RegPath,
        [string]$ExtensionId
    )
    try {
        if (-not (Test-Path $RegPath)) { return $false }
        $item = Get-Item $RegPath -ErrorAction SilentlyContinue
        if ($null -eq $item -or $null -eq $item.Property) { return $false }
        foreach ($name in $item.Property) {
            $val = (Get-ItemProperty -Path $RegPath -Name $name -ErrorAction SilentlyContinue).$name
            if ($val -like "$ExtensionId;*") { return $true }
        }
        return $false
    }
    catch { return $false }
}

function Get-NextForcelistIndex {
    param([string]$RegPath)
    if (-not (Test-Path $RegPath)) { return 1 }
    $props = (Get-Item $RegPath -ErrorAction SilentlyContinue).Property
    if ($null -eq $props -or $props.Count -eq 0) { return 1 }
    $nums = @($props | ForEach-Object { [int]$_ } | Sort-Object)
    return ($nums[-1] + 1)
}


# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "browser_addon_list.csv"

$enabledItems = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Browser", "ExtensionId", "Description")

if ($null -eq $enabledItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load browser_addon_list.csv")
}
if ($enabledItems.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}


# ========================================
# Step 2: 前処理（ID解決・Browser検証）
# ========================================
$resolvedItems = @()

foreach ($item in $enabledItems) {
    $entry = $item | Select-Object *
    $browser = $item.Browser.Trim()

    # Browser validation
    $regPath = switch ($browser) {
        'Chrome' { 'HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist' }
        'Edge'   { 'HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist' }
        default  { $null }
    }

    if ($null -eq $regPath) {
        $entry | Add-Member -NotePropertyName 'ResolvedId' -NotePropertyValue $null
        $entry | Add-Member -NotePropertyName 'RegPath'    -NotePropertyValue $null
        $entry | Add-Member -NotePropertyName 'IsInvalid'  -NotePropertyValue $true
        $entry | Add-Member -NotePropertyName 'ErrorReason' -NotePropertyValue "Unsupported browser: $browser (Chrome or Edge only)"
        $resolvedItems += $entry
        continue
    }

    # Resolve extension ID
    $resolvedId = Resolve-ExtensionId -InputValue $item.ExtensionId

    if ($null -eq $resolvedId) {
        $entry | Add-Member -NotePropertyName 'ResolvedId' -NotePropertyValue $null
        $entry | Add-Member -NotePropertyName 'RegPath'    -NotePropertyValue $regPath
        $entry | Add-Member -NotePropertyName 'IsInvalid'  -NotePropertyValue $true
        $entry | Add-Member -NotePropertyName 'ErrorReason' -NotePropertyValue "Cannot resolve extension ID from: $($item.ExtensionId)"
        $resolvedItems += $entry
        continue
    }

    $entry | Add-Member -NotePropertyName 'ResolvedId'  -NotePropertyValue $resolvedId
    $entry | Add-Member -NotePropertyName 'RegPath'     -NotePropertyValue $regPath
    $entry | Add-Member -NotePropertyName 'IsInvalid'   -NotePropertyValue $false
    $entry | Add-Member -NotePropertyName 'ErrorReason' -NotePropertyValue $null
    $resolvedItems += $entry
}


# ========================================
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Target Extensions" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

foreach ($item in $resolvedItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.ExtensionId }

    if ($item.IsInvalid) {
        Write-Host "  [ERROR] $displayName" -ForegroundColor Red
        Write-Host "    Browser:      $($item.Browser)" -ForegroundColor DarkGray
        Write-Host "    Reason:       $($item.ErrorReason)" -ForegroundColor Red
        Write-Host ""
        continue
    }

    $isRegistered = Test-ExtensionInForcelist -RegPath $item.RegPath -ExtensionId $item.ResolvedId
    $marker = if ($isRegistered) { "[Current]" } else { "[Change]" }
    $markerColor = if ($isRegistered) { "Gray" } else { "White" }

    Write-Host "  $marker $displayName" -ForegroundColor $markerColor
    Write-Host "    Browser:      $($item.Browser)" -ForegroundColor DarkGray
    Write-Host "    Extension ID: $($item.ResolvedId)" -ForegroundColor DarkGray
    Write-Host "    Registry:     $($item.RegPath)" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Apply the above browser extension policies?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: 設定適用ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($item in $resolvedItems) {
    $displayName = if ($item.Description) { $item.Description } else { $item.ResolvedId }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Processing: $displayName" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # Invalid entry
    if ($item.IsInvalid) {
        Show-Error "$($item.ErrorReason)"
        $failCount++
        Write-Host ""
        continue
    }

    # Idempotency check
    if (Test-ExtensionInForcelist -RegPath $item.RegPath -ExtensionId $item.ResolvedId) {
        Show-Skip "Already in forcelist"
        $skipCount++
        Write-Host ""
        continue
    }

    try {
        $value = "$($item.ResolvedId);https://clients2.google.com/service/update2/crx"

        # Create registry key if not exists
        if (-not (Test-Path $item.RegPath)) {
            Write-Host "  -> Creating registry key: $($item.RegPath)" -ForegroundColor Gray
            New-Item -Path $item.RegPath -Force | Out-Null
        }

        # Get next index
        $nextIndex = Get-NextForcelistIndex -RegPath $item.RegPath

        # Write the entry
        New-ItemProperty -Path $item.RegPath -Name $nextIndex -Value $value `
            -PropertyType String -Force -ErrorAction Stop | Out-Null

        Show-Success "Registered as entry #$nextIndex ($($item.Browser))"
        $successCount++
    }
    catch {
        Show-Error "Failed: $_"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Browser Addon Configuration Results")
