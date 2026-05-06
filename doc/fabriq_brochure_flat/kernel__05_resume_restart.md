# 再起動跨ぎ・Resume 機構

fabriq の最も特徴的な仕組み。`__RESTART__` マーカー 1 つで、マシン再起動 → 自動ログオン → fabriq 自動再起動 → 環境変数復元 → 残モジュール継続実行までを成立させる。

---

## __RESTART__ の制御フロー

```
profile CSV 中に __RESTART__ 行を検出
   ↓
Invoke-BatchExecution の loop が _IsRestart=true を見つける
   ↓
1. Save-ResumeState
   ├── 現セッションの SELECTED_* 環境変数を全部 hash table 化
   ├── ProfilePath / ProfileName / SessionID / ResumeAfterOrder（restart 行の Order）
   ├── CompletedModules（ここまで終わった module の Order/MenuName/Status 配列）
   ├── AutoPilot / AutoPilotWaitSec
   ├── EvidenceBasePath（再 init せずに同じ evidence dir を使い続ける）
   ├── ProfileStartTime（ISO 8601）── 経過時間を gap 込みで計算するため絶対起点を保持
   ├── ProtectedPassphrase（DPAPI LocalMachine で暗号化したマスターパスフレーズ）
   ├── FlexProfile の場合: schemaVersion=2 + ExecutionMode='Flex' + SelectedOrders + ModuleStates
   └── kernel/json/resume_state.json に書き込み
   ↓
2. Register-FabriqRunOnce
   ├── HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\FabriqAutoStart
   └── 値: "<fabriqRoot>\Fabriq.exe"
   （RunOnce は次回起動時に 1 回だけ実行されて自動消滅）
   ↓
3. Add-ExecutionResult [RESTART] Success / Write-ExecutionHistory
   ↓
4. Invoke-CountdownRestart -Seconds 5
   ├── 5 秒カウントダウン（Ctrl+C で abort）
   └── Restart-Computer -Force
   ↓ 物理的な再起動 ↓
   ↓
   AutoLogon が autologon_config 由来で組まれていれば自動ログオン
   ↓
   RunOnce による Fabriq.exe 自動起動 → main.ps1 冒頭
   ↓
5. main.ps1 が Load-ResumeState で resume_state.json を見つける
   ├── schemaVersion=2 + ExecutionMode='Flex' → FlexProfile resume ルート
   ├── AutoPilot=true → Wait-SystemReady（services 起動待ち）+ Invoke-AutoResumeCountdown
   └── 通常 → Confirm-Execution "Resume profile execution?"
   ↓
6. 環境復元
   ├── Restore-HostEnvironment（SELECTED_* env vars を json から戻す）
   ├── EvidenceBasePath 復元（or 再 init）
   ├── DPAPI で master passphrase 復元（失敗時は手動再入力 3 回まで）
   └── SessionID 復元（同一セッション扱いで履歴連結）
   ↓
7. Invoke-BatchExecution に再入し、CompletedModules の Order より大きい Order の
   モジュールから続行
   ├── Linear: profile 末尾まで自動進行
   └── Flex: dashboard 再 open で operator が次行動を選択
   ↓
8. 完走で Remove-ResumeState
```

---

## resume_state.json スキーマ

### v1 スキーマ（Linear、kernel 2.0.0〜）

```json
{
  "ProfilePath": "C:\\fabriq\\profiles\\Master_Pre01.csv",
  "ProfileName": "Master_Pre01",
  "AutoPilot": true,
  "AutoPilotWaitSec": 3,
  "SessionID": "20260506_103045",
  "ResumeAfterOrder": 40,
  "CompletedModules": [
    { "Order": 10, "MenuName": "Hostname Configuration", "Status": "Success" },
    { "Order": 20, "MenuName": "IP Address Configuration", "Status": "Success" },
    { "Order": 30, "MenuName": "Domain Join", "Status": "Partial" }
  ],
  "HostEnvironment": {
    "SELECTED_KANRI_NO": "1",
    "SELECTED_NEW_PCNAME": "NEW-PC-01",
    "SELECTED_ETH_IP": "192.168.1.100",
    /* ...全 SELECTED_* + SELECTED_PRINTER_*..* ... */
  },
  "EvidenceBasePath": ".\\evidence\\2026_05_06_103045_NEW-PC-01_T2NXCV06Y22208C_evidence\\evidence",
  "ProfileStartTime": "2026-05-06T10:30:45.1234567+09:00",
  "ProtectedPassphrase": "AQAAA...DPAPI Base64..."
}
```

`schemaVersion` フィールドは存在しない（Linear のシグネチャ）。

### v2 スキーマ（FlexProfile、kernel 3.1.0〜）

v1 のフィールドに加えて：

```json
{
  "schemaVersion": 2,
  "ExecutionMode": "Flex",
  "SelectedOrders": [10, 20, 40, 50],         // チェックボックスで選んだ Order 群
  "ModuleStates": {                            // 各 Order の現在状態（dashboard が再現する用）
    "10": { "Status": "Success",  "Verified": "True",  "Message": "Done" },
    "20": { "Status": "Success",  "Verified": "True",  "Message": "Done" },
    "30": { "Status": "Pending",  "Verified": "",      "Message": "" },
    "40": { "Status": "Error",    "Verified": "False", "Message": "Network timeout" }
  }
}
```

#### ResumeAfterOrder=-1 sentinel

FlexProfile の `[Restart Now]` ボタンで再起動した場合の特殊値。Linear-style の auto-continuation を発動させずに「dashboard を再 open するだけ」の挙動になる。post-reboot 後、operator が手動で次に何を回すか選べる。

### 後方互換ルール

- v1 file は v2 reader で問題なく読める（`schemaVersion` 不在 → v1 として扱う）
- v2 file は v1 reader（kernel 3.0.x）で読むと `SelectedOrders` 等が無視されるだけで Linear resume として機能する（graceful degradation）

---

## RunOnce 登録の実装

```powershell
function Register-FabriqRunOnce {
    $fabriqExe = Join-Path (Resolve-Path ".").Path "Fabriq.exe"
    $runOncePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    New-ItemProperty -Path $runOncePath -Name "FabriqAutoStart" `
        -Value "`"$fabriqExe`"" -PropertyType String -Force
}
```

- HKLM の RunOnce → **すべてのユーザーログオン**で 1 回だけ発火（HKCU だと特定ユーザーログオン時のみ）
- 値は引用符で囲んだ Fabriq.exe 絶対パス
- 起動した Fabriq.exe は内部で UAC 昇格を再度要求 → 管理者権限で main.ps1 が走る

---

## Wait-SystemReady（再起動直後の安定待ち）

Windows の再起動直後はサービスが起動中で、即座にモジュールを動かすと失敗するため、resume 経路では必ずこの関数を通る：

```powershell
Wait-SystemReady -MaxWaitSec 120 `
                 -RequiredServices @("LanmanWorkstation", "Dnscache") `
                 [-NetworkTarget "8.8.8.8"]
```

- 指定サービスがすべて Running になるまで 5 秒間隔で polling
- NetworkTarget 指定時は `Test-Connection` も追加判定
- MaxWaitSec を超えたら `Show-Warning` 出して続行（ハング防止）

---

## Invoke-AutoResumeCountdown（AutoPilot 再開時の確認）

```powershell
Invoke-AutoResumeCountdown -Seconds 60
   → returns $true  : Resume execution
   → returns $false : Abort (clear resume state)
```

UI:

```
[AUTOPILOT] Auto-resume countdown
  [Enter] = Resume now   [Esc] = Abort

  Resuming in 47 seconds...
```

- `Enter` / `Y` で即時 resume
- `Esc` / `N` / `Q` で中断 → `Remove-ResumeState`
- 何も押さなければカウントダウン消化後 auto-resume

「無人キッティング」の `完全自動` ではない。AutoPilot は「**確認スキップ + auto-resume**」であり、operator は脇で見ていて状況に応じて Esc できる前提（feedback memory `feedback_autopilot_wording`）。

---

## Reset-FabriqState（同一プロセスでの新セッション開始）

`Refabriq` ボタン / `New Session` で発火する完全リセット：

1. Transcript 停止 → 新規 transcript 開始（新ファイル名）
2. `$script:ExecutionResults` / `$script:LastBatchResults` / `$script:SessionID` をリセット
3. `execution_history.csv` および `.bak` を削除（次セッションを clean 状態に）
4. `session.json` 削除（worker 再選択を強制）
5. `$global:AutoPilotMode = $false` / `_LastModuleResult = $null`
6. `$global:FabriqEvidenceBasePath` / Root Path / `FABRIQ_EVIDENCE_BASE` を null 化
7. `SELECTED_*` 環境変数を全削除
8. `Remove-ResumeState`
9. `Write-StatusFile -Phase "idle"`

evidence ディレクトリは disk から消さない（前セッションの成果物を保護）。
