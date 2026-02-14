# 設定ファイルのパス
$csvPath = "commands\recipe.csv"

# .NETのSendKeysを利用（安定性のため）
Add-Type -AssemblyName System.Windows.Forms

# CSVを読み込む
$commands = Import-Csv $csvPath -Encoding Default

# 1行ずつ実行するループ
foreach ($row in $commands) {
    Write-Host "実行中: $($row.Note) - Action: $($row.Action)" -ForegroundColor Cyan

    # 動作ごとの処理分岐（ここがエンジンの核）
    switch ($row.Action) {
        "Open" {
            # アプリを起動する
            Start-Process $row.Value
        }
        "AppFocus" {
            # 特定のウィンドウを最前面にする (VBSのAppActivate相当)
            $ws = New-Object -ComObject WScript.Shell
            $ws.AppActivate($row.Value)
        }
        "Type" {
            # 文字列を入力する
            [System.Windows.Forms.SendKeys]::SendWait($row.Value)
        }
        "Key" {
            # 特殊キー（EnterやTabなど）を送信する
            [System.Windows.Forms.SendKeys]::SendWait($row.Value)
        }
        "Wait" {
            # 追加で待機が必要な場合
            Start-Sleep -Milliseconds ([int]$row.Value)
        }
        Default {
            Write-Warning "未定義のアクションです: $($row.Action)"
        }
    }

    # 各ステップごとの待機（CSVのWait列を参照）
    if ($row.Wait -gt 0) {
        Start-Sleep -Milliseconds ([int]$row.Wait)
    }
}

Write-Host "すべての処理が完了しました。" -ForegroundColor Green