<#
.SYNOPSIS
    Windows Comprehensive History Cleaner
.DESCRIPTION
    エクスプローラ、検索、イベントログ、Office、IME予測変換(ファイルベース)、
    ジャンプリストなどを一括削除します。
#>

# 管理者権限チェック
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
    Write-Warning "このスクリプトは管理者権限で実行する必要があります。右クリックして『管理者として実行』してください。"
    Pause
    Exit
}

# エラーを無視して続行（使用中ファイルなどで止まらないようにする）
$ErrorActionPreference = "SilentlyContinue"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   Windows 全履歴 強力削除プロセス開始" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 1. エクスプローラープロセスの停止 (ファイルのロックを解除するため)
Write-Host "[1/8] エクスプローラーを停止中..." -ForegroundColor Yellow
Stop-Process -Name "explorer" -Force

# 2. エクスプローラー関連の履歴削除
Write-Host "[2/8] エクスプローラー履歴(MRU/Run/検索)を削除中..." -ForegroundColor Yellow

# 最近使ったファイル (Recent)
Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\*" -Recurse -Force
Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\*" -Recurse -Force # ジャンプリスト
Remove-Item "$env:APPDATA\Microsoft\Windows\Recent\CustomDestinations\*" -Recurse -Force

# レジストリ: 最近使ったファイル/検索履歴/ファイル名を指定して実行
$ExplorerPaths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\WordWheelQuery", # エクスプローラ検索ボックス履歴
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StreamMRU"
)

foreach ($path in $ExplorerPaths) {
    if (Test-Path $path) {
        Remove-ItemProperty -Path $path -Name * -ErrorAction SilentlyContinue
    }
}

# 3. イベントビューア（全ログ消去）
Write-Host "[3/8] イベントビューアの全ログを消去中..." -ForegroundColor Yellow
# wevtutilを使用して全ログリストを取得しクリアする
Get-WinEvent -ListLog * -Force | ForEach-Object { 
    wevtutil.exe cl $_.LogName 
}

# 4. Microsoft Office履歴 (レジストリMRU)
Write-Host "[4/8] Office (Word/Excel等) の最近使ったファイルを削除中..." -ForegroundColor Yellow
# 各バージョンのOfficeレジストリを探索 (11.0 ~ 16.0など)
$OfficeBase = "HKCU:\Software\Microsoft\Office"
if (Test-Path $OfficeBase) {
    Get-ChildItem $OfficeBase | ForEach-Object {
        $Version = $_.PSChildName
        $Apps = @("Word", "Excel", "PowerPoint", "Access", "Publisher", "Visio")
        foreach ($App in $Apps) {
            $PlaceMRU = "$OfficeBase\$Version\$App\Place MRU"
            $FileMRU  = "$OfficeBase\$Version\$App\File MRU"
            
            if (Test-Path $PlaceMRU) { Remove-ItemProperty -Path $PlaceMRU -Name * -ErrorAction SilentlyContinue }
            if (Test-Path $FileMRU) { Remove-ItemProperty -Path $FileMRU -Name * -ErrorAction SilentlyContinue }
        }
    }
}

# 5. IME / 変換履歴 (ユーザー辞書キャッシュ)
Write-Host "[5/8] IME変換履歴キャッシュを削除中..." -ForegroundColor Yellow
# 注: クラウド候補やシステムメモリ上のものは完全に消えない場合があります
$ImePath = "$env:APPDATA\Microsoft\InputMethod"
if (Test-Path $ImePath) {
    Remove-Item "$ImePath\*" -Recurse -Force
}

# 6. 一時ファイル、クリップボード、DNS
Write-Host "[6/8] 一時ファイル・クリップボード・DNSをクリア中..." -ForegroundColor Yellow
Remove-Item "$env:TEMP\*" -Recurse -Force
Remove-Item "$env:windir\Temp\*" -Recurse -Force
Set-Clipboard $null # クリップボードを空にする
Clear-DnsClientCache # DNSキャッシュクリア

# 7. ごみ箱を空にする (C# API利用)
Write-Host "[7/8] ごみ箱を完全に空にしています..." -ForegroundColor Yellow
$Code = @'
    [DllImport("Shell32.dll")]
    public static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, int dwFlags);
'@
Add-Type -MemberDefinition $Code -Name Win32 -Namespace Native
# フラグ: SHERB_NOCONFIRMATION = 0x00000001, SHERB_NOPROGRESSUI = 0x00000002, SHERB_NOSOUND = 0x00000004
[Native.Win32]::SHEmptyRecycleBin([IntPtr]::Zero, $null, 7)

# 8. 終了処理
Write-Host "[8/8] エクスプローラーを再起動します..." -ForegroundColor Green
Start-Process "explorer.exe"

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "   すべての履歴削除が完了しました。" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Start-Sleep -Seconds 3