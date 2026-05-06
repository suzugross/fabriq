# ステータスモニタ + 演出 (ART pulse / silence / manifesto)

fabriq の二画面構成の左右の片方。メインダッシュボードとは独立した別プロセスで動き、`status.json` ファイルを polling する設計。これにより重い WinForms 処理がメイン実行をブロックしない。

---

## アーキテクチャ概観

```
[ Main Process: Fabriq.exe → main.ps1 ]
   ├── Show-* / Add-ExecutionResult が呼ばれる度に
   │   Write-StatusFile -Phase "executing"  ── status.json を atomic 書き込み
   │   Write-ArtPulse                          ── art_pulse.txt のカウンタを +1
   ↓ ファイルベース IPC ↓
[ Monitor Process: kernel/ps1/status_monitor.ps1（別 PowerShell プロセス）]
   ├── status.json を 500ms 間隔で polling
   ├── 変更検出時に WinForms TextBox / Label / Grid を更新
   └── 演出:
       ├── art_pulse.txt の数字が増えたら「鼓動」アニメ
       ├── art_sentences.txt から 1 行をランダム表示
       └── silence.flag が存在すれば演出を完全停止
```

---

## status.json スキーマ

`Write-StatusFile -Phase <idle|executing|complete>` が atomic write する：

```json
{
  "UpdatedAt": "2026-05-06 10:32:15",
  "WorkerName": "suzuki",
  "PCInfo": {
    "AdminID": "1",
    "OldPCName": "OLD-PC-01",
    "NewPCName": "NEW-PC-01",
    "EthernetIP": "192.168.1.100",
    "EthernetSubnet": "255.255.255.0",
    "EthernetGateway": "192.168.1.1",
    "WifiIP": "",
    "WifiSubnet": "",
    "WifiGateway": "",
    "DNS": ["8.8.8.8", "8.8.4.4"],
    "Printers": [
      { "Name": "Office", "Driver": "Canon Generic Plus PCL6", "Port": "192.168.1.50" }
    ]
  },
  "CurrentPCInfo": {
    /* Get-CurrentPCInfo の結果 — 実 OS から取った現在値（照合用） */
    "ComputerName": "NEW-PC-01",
    "EthernetIP": "192.168.1.100",
    /* ... */
    "Printers": [{ "Name": "Office", "Port": "192.168.1.50" }]
  },
  "Execution": {
    "Phase": "executing",
    "TotalCount": 12,
    "SuccessCount": 9,
    "ErrorCount": 0,
    "SkippedCount": 1,
    "CancelledCount": 0,
    "PartialCount": 1,
    "WarningCount": 0,
    "Details": [
      {
        "Operation": "Hostname Configuration",
        "Status": "Success",
        "Message": "Renamed to NEW-PC-01 (pending reboot)",
        "Timestamp": "2026-05-06 10:30:55",
        "IsRestored": false,
        "Verified": true
      },
      /* ... */
    ]
  }
}
```

### Atomic Write

```powershell
$tempPath = "$($script:StatusFilePath).tmp"
$statusData | ConvertTo-Json -Depth 5 | Out-File -FilePath $tempPath -Encoding UTF8 -Force
Move-Item -Path $tempPath -Destination $script:StatusFilePath -Force
```

書き込み中の半端な JSON を monitor が読んで crash しないよう、tmp ファイルへ書いてから rename。失敗時は直接書きへフォールバック（best-effort で例外は飲む）。

### PCInfo vs CurrentPCInfo

- `PCInfo`: hostlist.csv 起源の **期待値**（SELECTED_* env vars）
- `CurrentPCInfo`: `Get-CurrentPCInfo` が `Get-NetAdapter` / `Get-Printer` で取った **現在値**

monitor 画面の左右に並べて差分ハイライト。HTML チェックリストでも同じ照合を行う。

---

## status_monitor.ps1 の構造

別プロセスで起動される独立 WinForms アプリ。

### 起動コマンド

```powershell
Start-Process powershell.exe -ArgumentList @(
    "-NoProfile", "-ExecutionPolicy", "Unrestricted",
    "-File", ".\kernel\ps1\status_monitor.ps1",
    "-StatusFilePath", $statusFileFullPath,
    "-PulseFilePath", $pulseFileFullPath,
    "-SentenceFilePath", $sentenceFileFullPath,
    "-SilenceFlagPath", $silenceFlagFullPath
) -WindowStyle Hidden -PassThru
```

PID は `$global:FabriqStatusMonitorProcess` に格納し、`Stop-StatusMonitor` で `CloseMainWindow` → 2 秒待ち → 強制 Kill。

### 表示要素（概念）

- **PC Info ペイン**: PCInfo（期待値）と CurrentPCInfo（実際値）を縦並び
- **Execution Summary**: TotalCount / Success / Error / Skipped / Partial の数字
- **Execution Details Grid**: Operation, Status（色付）, Verified バッジ, Timestamp
- **Skip ボタン**: async モジュール強制中断（`skip_request.flag` を生成）
- **Manual Screenshot ボタン**: `Save-Screenshot` を呼んで `evidence/{base}/gyotaku/` に PNG 保存
- **ART Pulse 鼓動**: `art_pulse.txt` の数字が増えるたびにアニメ
- **ART 言葉**: `art_sentences.txt` から 1 行をランダム表示

### Skip ボタンの実装

monitor が `skip_request.flag` を作成 → `Invoke-SafeCommandAsync` の polling loop が検出 → Runspace を `ps.Stop()` で強制終了 → Skip 結果として記録。詳細は §08_async_execution.md。

---

## ART Pulse / Sentences / silence.flag（演出機能）

fabriq には「**Manifeste du Surkitinisme**」という演出文化がある（README L1）。Status Monitor がこの演出を担う。

### art_pulse.txt

- 中身: 整数 1 行（カウンタ）
- 書き手: `Write-ArtPulse`（`Show-Info` / `Show-Success` / `Show-Warning` / `Show-Error` / `Show-Skip` のすべてが内部で呼ぶ）
- 読み手: `status_monitor.ps1` が polling し、増加分を「鼓動」アニメに変換
- 用途: モジュール実行が「生きている」ことを視覚化（プロセスがハングしてないか operator が一目で判断できる）

### art_sentences.txt

- 中身: 演出用の 1 行テキスト集（1 sentence per line）
- 表示: monitor が定期的にランダム選択して大文字で表示
- 中身は内輪文化を反映（例: "DEKITINISME"）

### silence.flag

- 存在するだけで monitor が ART 演出を完全停止
- 業務環境で演出が邪魔な場合の opt-out 機構
- ファイル中身は問わない（存在チェックのみ）

---

## manifesto.ps1（マニフェスト表示 GUI）

`kernel/ps1/manifesto.ps1` に WinForms 関数 `Show-Manifesto` を定義。`kernel/csv/manifesto.csv`（章ごとの本文）を読み、ダッシュボードの `[Manifeste du Surkitinisme]` ボタンから呼ばれる。

純粋に演出機能であり、運用上の意味は無い（fabriq の哲学を作業者に表示する）。

---

## view_report.ps1（HTML チェックリストビューア）

`kernel/ps1/view_report.ps1` は単体起動可能な HTML ビューアスクリプト。プロファイル完了時に `Complete-ProfileExecution` から自動起動され、最新の `evidence/{base}/checklist/checklist_*.html` を既定ブラウザで開く。

---

## Status Monitor のライフサイクル

```
fabriq 起動
   ↓
Initialize-Session 後
   ↓
Start-StatusMonitor
   ├── Write-StatusFile -Phase "idle" でスケルトン生成
   ├── Start-Process powershell.exe ... -WindowStyle Hidden -PassThru
   ├── PID を $global:FabriqStatusMonitorProcess に保存
   └── 1.2 秒待ってから Set-ConsoleForeground（メインコンソールを前面に戻す）
   ↓
... モジュール実行中、Add-ExecutionResult 等のたびに Write-StatusFile が更新 ...
   ↓
Exit-Fabriq（ユーザーが終了）
   ↓
Stop-StatusMonitor
   ├── CloseMainWindow → 2 秒で Wait→ 強制 Kill
   └── Remove-StatusFile（status.json + tmp + art_pulse.txt 削除）
```
