# 1. ユーザーへの確認
Add-Type -AssemblyName System.Windows.Forms
$result = [System.Windows.Forms.MessageBox]::Show("バックアップからデスクトップアイコンの配置レジストリを復元しますか？`n`n※反映には次回のサインインまたはPC再起動が必要です。", "復元の確認", "YesNo")

if ($result -ne "Yes") {
    Write-Host "キャンセルされました。" -ForegroundColor Yellow
    exit
}

# 2. 最新のバックアップファイルを探す
$backupDir = Join-Path $PSScriptRoot "backup"

if (!(Test-Path $backupDir)) {
    [System.Windows.Forms.MessageBox]::Show("backupフォルダが見つかりません。", "エラー")
    exit
}

# backupフォルダ内で最も新しい .reg ファイルを取得
$latestBackup = Get-ChildItem -Path $backupDir -Filter "DesktopBags_Backup_*.reg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($null -eq $latestBackup) {
    [System.Windows.Forms.MessageBox]::Show("復元可能なバックアップファイルが見つかりません。", "エラー")
    exit
}

Write-Host "使用するファイル: $($latestBackup.FullName)" -ForegroundColor Cyan

# 3. インポートの実行
try {
    # reg.exe を使用してインポート
    # /s オプションで「結合しました」というOS標準のダイアログを抑制します
    $process = Start-Process reg.exe -ArgumentList "import `"$($latestBackup.FullName)`"" -Wait -PassThru -NoNewWindow
    
    if ($process.ExitCode -eq 0) {
        # 4. 完了通知
        Write-Host "---" -ForegroundColor Green
        Write-Host "レジストリのインポートが完了しました。"
        Write-Host "反映させるには、PCを再起動するかサインアウトしてください。"
        Write-Host "---"
        [System.Windows.Forms.MessageBox]::Show("復元（レジストリ結合）が完了しました。`n`n設定を反映させるために、PCを再起動するかサインアウトしてください。", "完了")
    } else {
        throw "reg.exe がエラーコード $($process.ExitCode) を返しました。"
    }
}
catch {
    Write-Error "復元中にエラーが発生しました: $_"
    [System.Windows.Forms.MessageBox]::Show("エラーが発生しました。`n$($_.Exception.Message)", "エラー")
}