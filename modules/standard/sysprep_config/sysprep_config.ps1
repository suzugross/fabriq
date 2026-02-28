# ========================================
# Sysprep Config - Sysprep Preparation & Execution
# ========================================
# Generates unattend.xml and SetupComplete.cmd from CSV,
# deploys them to the system, and executes sysprep.
#
# [NOTES]
# - Requires administrator privileges
# - Sysprep execution is irreversible (shutdown/reboot)
# - Two-phase confirmation for safety
# ========================================

Write-Host ""
Show-Separator
Write-Host "Sysprep Config" -ForegroundColor Cyan
Show-Separator
Write-Host ""

# ========================================
# デプロイ先パス定義
# ========================================
$unattendDeployPath      = "C:\Windows\System32\Sysprep\unattend.xml"
$setupCompleteDeployDir  = "C:\Windows\Setup\Scripts"
$setupCompleteDeployPath = Join-Path $setupCompleteDeployDir "SetupComplete.cmd"
$sourceStagingDir        = Join-Path $setupCompleteDeployDir "source"
$sourceDir               = Join-Path $PSScriptRoot "source"

# ========================================
# Unattend.xml テンプレート（固定部分 + プレースホルダ）
# ========================================
# 固定部分:
#   generalize  → ドライバ保持（PersistAllDeviceInstalls 等）
#   specialize  → タイムゾーン（Tokyo Standard Time）
#   oobeSystem  → 言語設定（ja-JP）、デバイス暗号化防止
#
# 可変部分（プレースホルダ）:
#   {{SPECIALIZE_SETTINGS}} → ComputerName, CopyProfile
#   {{OOBE_BLOCK}}          → <OOBE> 内の各スキップ設定
#   {{USER_ACCOUNTS_BLOCK}} → TestUserName, EnableAdministrator
# ========================================
$xmlTemplate = @'
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
    <settings pass="generalize">
        <component name="Microsoft-Windows-PnpSysprep" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <DoNotCleanUpNonPresentDevices>true</DoNotCleanUpNonPresentDevices>
            <PersistAllDeviceInstalls>true</PersistAllDeviceInstalls>
        </component>
    </settings>
    <settings pass="specialize">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
{{SPECIALIZE_SETTINGS}}            <TimeZone>Tokyo Standard Time</TimeZone>
        </component>
    </settings>
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <InputLocale>ja-JP</InputLocale>
            <SystemLocale>ja-JP</SystemLocale>
            <UILanguage>ja-JP</UILanguage>
            <UILanguageFallback>ja-JP</UILanguageFallback>
            <UserLocale>ja-JP</UserLocale>
        </component>
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
{{OOBE_BLOCK}}{{USER_ACCOUNTS_BLOCK}}        </component>
        <component name="Microsoft-Windows-SecureStartup-FilterDriver" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <PreventDeviceEncryption>true</PreventDeviceEncryption>
        </component>
    </settings>
    <cpi:offlineImage cpi:source="wim:c:/install.wim#Windows 11 Pro" xmlns:cpi="urn:schemas-microsoft-com:cpi" />
</unattend>
'@

# ========================================
# SetupComplete.cmd テンプレート（ヘッダー・フッター）
# ========================================
$cmdHeader = @'
@echo off
echo [%date% %time%] SetupComplete.cmd execution started > C:\Windows\Setup\Scripts\SetupComplete.log
'@

$cmdFooter = @'
echo [%date% %time%] SetupComplete.cmd execution completed >> C:\Windows\Setup\Scripts\SetupComplete.log
exit
'@

# OOBE 設定名の定義（カテゴリ判定用）
$oobeSettingNames = @(
    "HideEULAPage",
    "ProtectYourPC",
    "HideWirelessSetupInOOBE",
    "HideOnlineAccountScreens",
    "HideOEMRegistrationScreen"
)


# ========================================
# Step 1: CSV 読み込み（3 ファイル）
# ========================================

# --- 1-1: sysprep_list.csv（sysprep 実行設定 / 単一行） ---
$sysprepCsvPath = Join-Path $PSScriptRoot "sysprep_list.csv"

$sysprepItems = Import-ModuleCsv -Path $sysprepCsvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SysprepExe", "Mode", "Shutdown")

if ($null -eq $sysprepItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load sysprep_list.csv")
}
if ($sysprepItems.Count -eq 0) {
    return (New-ModuleResult -Status "Error" -Message "No enabled sysprep configuration")
}

$sysprepConfig = $sysprepItems[0]

# --- 1-2: unattend_list.csv（応答ファイル設定 / キー・バリュー） ---
$unattendCsvPath = Join-Path $PSScriptRoot "unattend_list.csv"

$unattendItems = Import-ModuleCsv -Path $unattendCsvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "SettingName", "Value")

if ($null -eq $unattendItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load unattend_list.csv")
}

# キー・バリュー形式のハッシュテーブルに変換（0 件でも続行 = 固定値のみの XML を生成）
$settings = @{}
foreach ($item in $unattendItems) {
    $settings[$item.SettingName] = $item.Value
}

# --- 1-3: setupcomplete_list.csv（SetupComplete アクション定義） ---
$setupCsvPath = Join-Path $PSScriptRoot "setupcomplete_list.csv"

$setupItems = Import-ModuleCsv -Path $setupCsvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Order", "ActionType", "Target")

if ($null -eq $setupItems) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load setupcomplete_list.csv")
}

# Order 昇順ソート（0 件でも続行 = ヘッダー・フッターのみの CMD を生成）
if ($setupItems.Count -gt 0) {
    $setupItems = @($setupItems | Sort-Object { [int]$_.Order })
}

# CopyFile アクションの抽出
$copyFileActions = @($setupItems | Where-Object { $_.ActionType -eq "CopyFile" })


# ========================================
# Step 2: 前提条件チェック
# ========================================

# --- sysprep.exe の存在確認 ---
if (-not (Test-Path $sysprepConfig.SysprepExe)) {
    Show-Error "sysprep.exe not found: $($sysprepConfig.SysprepExe)"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "sysprep.exe not found")
}

# --- source/ ディレクトリの存在確認（CopyFile アクションがある場合のみ） ---
if ($copyFileActions.Count -gt 0) {
    if (-not (Test-Path $sourceDir)) {
        Show-Error "source/ directory not found: $sourceDir"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "source/ directory not found")
    }
}

# --- デプロイ先ディレクトリの作成 ---
foreach ($dir in @($setupCompleteDeployDir, $sourceStagingDir)) {
    if (-not (Test-Path $dir)) {
        try {
            $null = New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop
            Show-Info "Created directory: $dir"
        }
        catch {
            Show-Error "Failed to create directory: $dir - $_"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Deploy directory creation failed")
        }
    }
}


# ========================================
# Step 3: 実行前の確認表示
# ========================================

# --- [Sysprep 設定] ---
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Sysprep Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Executable: $($sysprepConfig.SysprepExe)" -ForegroundColor White
Write-Host "  Mode:       /$($sysprepConfig.Mode)" -ForegroundColor White
Write-Host "  Shutdown:   /$($sysprepConfig.Shutdown)" -ForegroundColor White
Write-Host ""

# --- [Unattend.xml 設定一覧] ---
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Unattend.xml Settings" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($unattendItems.Count -eq 0) {
    Write-Host "  (defaults only)" -ForegroundColor DarkGray
}
else {
    foreach ($item in $unattendItems) {
        # AdminPassword はマスク表示
        $displayValue = if ($item.SettingName -eq "AdminPassword" -and $item.Value -ne "") {
            "********"
        } else {
            $item.Value
        }
        Write-Host "  $($item.SettingName) = $displayValue" -ForegroundColor White
    }
}
Write-Host ""

# --- [SetupComplete.cmd アクション一覧] ---
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "SetupComplete.cmd Actions" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if ($setupItems.Count -eq 0) {
    Write-Host "  (no actions)" -ForegroundColor DarkGray
}
else {
    $index = 0
    foreach ($action in $setupItems) {
        $index++

        switch ($action.ActionType) {
            "DeleteUser" {
                Write-Host "  [$index] DeleteUser: $($action.Target)" -ForegroundColor White
            }
            "CopyFile" {
                $srcCheckPath = Join-Path $sourceDir $action.Target
                if (Test-Path $srcCheckPath) {
                    $marker = "[FOUND]"
                    $markerColor = "White"
                }
                else {
                    $marker = "[NOT FOUND]"
                    $markerColor = "Red"
                }
                Write-Host "  [$index] CopyFile: $($action.Target)  $marker" -ForegroundColor $markerColor
                Write-Host "      Dest: $($action.Destination)" -ForegroundColor DarkGray
            }
            "Command" {
                $cmdDisplay = if ($action.Target.Length -gt 70) {
                    $action.Target.Substring(0, 67) + "..."
                } else {
                    $action.Target
                }
                Write-Host "  [$index] Command: $cmdDisplay" -ForegroundColor White
            }
        }

        if ($action.Description) {
            Write-Host "      $($action.Description)" -ForegroundColor Gray
        }
        Write-Host ""
    }
}

# --- [デプロイ先] ---
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Deploy Targets" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""

if (Test-Path $unattendDeployPath) {
    Write-Host "  unattend.xml      -> $unattendDeployPath  [OVERWRITE]" -ForegroundColor Yellow
}
else {
    Write-Host "  unattend.xml      -> $unattendDeployPath  [NEW]" -ForegroundColor White
}

if (Test-Path $setupCompleteDeployPath) {
    Write-Host "  SetupComplete.cmd -> $setupCompleteDeployPath  [OVERWRITE]" -ForegroundColor Yellow
}
else {
    Write-Host "  SetupComplete.cmd -> $setupCompleteDeployPath  [NEW]" -ForegroundColor White
}
Write-Host ""


# ========================================
# Step 4: 第 1 確認（ファイル生成・配置）
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Generate and deploy configuration files?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: ファイル生成 & デプロイ
# ========================================

# --- 5-1: source/ → ステージングディレクトリへコピー ---
if (Test-Path $sourceDir) {
    $sourceContents = @(Get-ChildItem -Path $sourceDir -ErrorAction SilentlyContinue)
    if ($sourceContents.Count -gt 0) {
        try {
            $null = Copy-Item -Path "$sourceDir\*" -Destination $sourceStagingDir -Recurse -Force -ErrorAction Stop
            Show-Success "Source files staged: $($sourceContents.Count) items -> $sourceStagingDir"
        }
        catch {
            Show-Error "Failed to stage source files: $_"
            Write-Host ""
            return (New-ModuleResult -Status "Error" -Message "Source staging failed: $_")
        }
    }
}

# --- 5-2: unattend.xml 動的生成 → デプロイ ---
Show-Info "Generating unattend.xml..."

# {{SPECIALIZE_SETTINGS}} の組み立て
$specializeBlock = ""
if ($settings.ContainsKey("ComputerName")) {
    $specializeBlock += "            <ComputerName>$($settings["ComputerName"])</ComputerName>`r`n"
}
if ($settings.ContainsKey("CopyProfile")) {
    $specializeBlock += "            <CopyProfile>$($settings["CopyProfile"])</CopyProfile>`r`n"
}

# {{OOBE_BLOCK}} の組み立て
$oobeContent = ""
foreach ($name in $oobeSettingNames) {
    if ($settings.ContainsKey($name)) {
        $oobeContent += "                <$name>$($settings[$name])</$name>`r`n"
    }
}
if ($oobeContent -ne "") {
    $oobeBlock = "            <OOBE>`r`n$oobeContent            </OOBE>`r`n"
}
else {
    $oobeBlock = ""
}

# {{USER_ACCOUNTS_BLOCK}} の組み立て
$userAccountsContent = ""

# AdministratorPassword ブロック
if ($settings.ContainsKey("EnableAdministrator") -and $settings["EnableAdministrator"] -eq "true") {
    $adminPassword = ""
    if ($settings.ContainsKey("AdminPassword")) {
        $adminPassword = $settings["AdminPassword"]
    }
    $userAccountsContent += "                <AdministratorPassword>`r`n"
    $userAccountsContent += "                    <Value>$adminPassword</Value>`r`n"
    $userAccountsContent += "                    <PlainText>true</PlainText>`r`n"
    $userAccountsContent += "                </AdministratorPassword>`r`n"
}

# LocalAccount ブロック
if ($settings.ContainsKey("TestUserName") -and $settings["TestUserName"] -ne "") {
    $testUser = $settings["TestUserName"]
    $userAccountsContent += "                <LocalAccounts>`r`n"
    $userAccountsContent += "                    <LocalAccount wcm:action=`"add`">`r`n"
    $userAccountsContent += "                        <Name>$testUser</Name>`r`n"
    $userAccountsContent += "                        <Group>Administrators</Group>`r`n"
    $userAccountsContent += "                    </LocalAccount>`r`n"
    $userAccountsContent += "                </LocalAccounts>`r`n"
}

if ($userAccountsContent -ne "") {
    $userAccountsBlock = "            <UserAccounts>`r`n$userAccountsContent            </UserAccounts>`r`n"
}
else {
    $userAccountsBlock = ""
}

# プレースホルダ置換（.Replace で安全に文字列置換）
$xmlContent = $xmlTemplate.Replace('{{SPECIALIZE_SETTINGS}}', $specializeBlock)
$xmlContent = $xmlContent.Replace('{{OOBE_BLOCK}}', $oobeBlock)
$xmlContent = $xmlContent.Replace('{{USER_ACCOUNTS_BLOCK}}', $userAccountsBlock)

# unattend.xml 書き出し
try {
    $xmlContent | Out-File -FilePath $unattendDeployPath -Encoding UTF8 -Force -ErrorAction Stop

    if (-not (Test-Path $unattendDeployPath)) {
        Show-Error "unattend.xml was not created: $unattendDeployPath"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "unattend.xml not found after write")
    }

    $fileSize = (Get-Item $unattendDeployPath).Length
    if ($fileSize -eq 0) {
        Show-Error "unattend.xml is empty"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "unattend.xml is empty after write")
    }

    Show-Success "unattend.xml deployed ($fileSize bytes)"
    Write-Host "  Path: $unattendDeployPath" -ForegroundColor DarkGray
}
catch {
    Show-Error "Failed to write unattend.xml: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to write unattend.xml: $_")
}

# --- 5-3: SetupComplete.cmd 動的生成 → デプロイ ---
Show-Info "Generating SetupComplete.cmd..."

$cmdBody = $cmdHeader + "`r`n"

foreach ($action in $setupItems) {
    $cmdBody += "`r`n"
    $cmdBody += "REM --- [$($action.ActionType)] $($action.Target) ---`r`n"

    switch ($action.ActionType) {
        "DeleteUser" {
            $cmdBody += "net user `"$($action.Target)`" /delete`r`n"
        }
        "CopyFile" {
            $cmdBody += "xcopy /Y /E /C /I `"$sourceStagingDir\$($action.Target)`" `"$($action.Destination)`"`r`n"
        }
        "Command" {
            $cmdBody += "$($action.Target)`r`n"
        }
    }
}

$cmdBody += "`r`n"
$cmdBody += $cmdFooter

# SetupComplete.cmd 書き出し
try {
    $cmdBody | Out-File -FilePath $setupCompleteDeployPath -Encoding UTF8 -Force -ErrorAction Stop

    if (-not (Test-Path $setupCompleteDeployPath)) {
        Show-Error "SetupComplete.cmd was not created: $setupCompleteDeployPath"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "SetupComplete.cmd not found after write")
    }

    $fileSize = (Get-Item $setupCompleteDeployPath).Length
    if ($fileSize -eq 0) {
        Show-Error "SetupComplete.cmd is empty"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "SetupComplete.cmd is empty after write")
    }

    Show-Success "SetupComplete.cmd deployed ($fileSize bytes)"
    Write-Host "  Path: $setupCompleteDeployPath" -ForegroundColor DarkGray
}
catch {
    Show-Error "Failed to write SetupComplete.cmd: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Failed to write SetupComplete.cmd: $_")
}

Write-Host ""


# ========================================
# Step 5.5: 第 2 確認（Sysprep 実行）
# ========================================
$sysprepArgs = "/generalize /$($sysprepConfig.Mode) /$($sysprepConfig.Shutdown) /unattend:$unattendDeployPath"
$sysprepCommand = "$($sysprepConfig.SysprepExe) $sysprepArgs"

Write-Host "========================================" -ForegroundColor Red
Write-Host "Sysprep Execution" -ForegroundColor Red
Write-Host "========================================" -ForegroundColor Red
Write-Host ""
Write-Host "  Command: $sysprepCommand" -ForegroundColor White
Write-Host ""

Show-Warning "This operation is irreversible. The PC will $($sysprepConfig.Shutdown) after sysprep completes."
Write-Host ""

$cancelResult = Confirm-ModuleExecution -Message "Execute sysprep? (PC will $($sysprepConfig.Shutdown))"
if ($null -ne $cancelResult) {
    # ファイル配置は完了しているので Success で返す
    Show-Info "Files deployed successfully. Sysprep was not executed."
    Write-Host ""
    return (New-ModuleResult -Status "Success" -Message "Files deployed. Sysprep not executed.")
}

Write-Host ""

# --- Sysprep 実行 ---
Show-Info "Executing sysprep..."
try {
    $proc = Start-Process -FilePath $sysprepConfig.SysprepExe -ArgumentList $sysprepArgs `
        -Wait -PassThru -ErrorAction Stop

    # /quit モードの場合のみここに到達
    if ($proc.ExitCode -ne 0) {
        Show-Error "Sysprep exited with code: $($proc.ExitCode)"
        Write-Host ""
        return (New-ModuleResult -Status "Error" -Message "Sysprep failed (exit code: $($proc.ExitCode))")
    }

    Show-Success "Sysprep completed (exit code: $($proc.ExitCode))"
    Write-Host ""
}
catch {
    Show-Error "Sysprep execution failed: $_"
    Write-Host ""
    return (New-ModuleResult -Status "Error" -Message "Sysprep execution failed: $_")
}


# ========================================
# Step 6: 結果返却
# ========================================
# /quit モードの場合のみ到達
# /shutdown, /reboot の場合 → PC 停止のためここには到達しない（想定通り）
return (New-ModuleResult -Status "Success" -Message "Sysprep completed ($($sysprepConfig.Mode)/$($sysprepConfig.Shutdown))")
