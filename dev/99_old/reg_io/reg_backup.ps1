# 1. ユーザーへの通知と確認
Add-Type -AssemblyName System.Windows.Forms
$msg = "デスクトップアイコンを希望の並び順に整えてください。`n`n準備ができたら「OK」を押すとバックアップを開始します。"
[System.Windows.Forms.MessageBox]::Show($msg, "バックアップの準備")

# 2. 保存先ディレクトリの設定
# スクリプトがある場所を取得し、backupフォルダのパスを作成
$currentDir = $PSScriptRoot
$backupDir = Join-Path $currentDir "backup"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportPath = Join-Path $backupDir "DesktopBags_Backup_$($timestamp).reg"

# backupフォルダが存在しない場合は作成
if (!(Test-Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir | Out-Null
}

# 3. バックアップ（レジストリエクスポート）の実行
$registryPath = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Bags\1\Desktop'

try {
    # reg.exe を使用してエクスポート
    Start-Process reg.exe -ArgumentList "export `"$registryPath`" `"$exportPath`" /y" -Wait -NoNewWindow
    
    # 4. 処理完了の表示
    Write-Host "---" -ForegroundColor Cyan
    Write-Host "バックアップが完了しました！" -ForegroundColor Green
    Write-Host "保存先: $exportPath"
    Write-Host "---"
    
    [System.Windows.Forms.MessageBox]::Show("バックアップが完了しました。`n保存先: $exportPath", "完了")
}
catch {
    Write-Error "バックアップ中にエラーが発生しました: $_"
}