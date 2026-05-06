# common.ps1 関数インデックス（90+ 関数）

`kernel/common.ps1` で定義されている全関数の一覧。**公開 API**（KERNEL_API.md §1〜§5 で宣言）と**内部実装**（PATCH バージョンでも変更されうる）を区別して記載。

---

## A. 公開 API（モジュールから安全に依存可能）

### A.1 表示・通知（color-coded console output）

| 関数 | 行番号 | 用途 |
|---|---|---|
| `Show-Info` | 278 | シアン `[INFO]` |
| `Show-Success` | 284 | グリーン `[SUCCESS]` |
| `Show-Warning` | 290 | イエロー `[WARNING]` |
| `Show-Error` | 296 | レッド `[ERROR]` |
| `Show-Skip` | 302 | ダークグレー `[SKIP]` |
| `Show-Separator` | 256 | シアンの横線 |
| `Show-CategorySeparator` | 260 | `=== <Name> ===` |

### A.2 結果オブジェクト

| 関数 | 行番号 | 用途 |
|---|---|---|
| `New-ModuleResult` | 312 | ModuleResult 生成（_IsModuleResult, Status, Message, Details, Verified, Timestamp） |
| `New-BatchResult` | 342 | 集計表示 + 自動 Status 判定 + New-ModuleResult 呼び出し |
| `Confirm-ModuleExecution` | 388 | Y/N 確認（N で Cancelled の ModuleResult 返却） |

### A.3 CSV / 暗号化

| 関数 | 行番号 | 用途 |
|---|---|---|
| `Import-ModuleCsv` | 529 | CSV 読み込み + 透過復号 + 必須列検証 + Enabled / Segment フィルタ |
| `Unprotect-FabriqValue` | 445 | AES-256-CBC + PBKDF2-SHA256 復号 |

### A.4 ユーザー確認・待機

| 関数 | 行番号 | 用途 |
|---|---|---|
| `Confirm-Execution` | 625 | Y/N → bool（AutoPilot/AutoConfirm 自動 Y） |
| `Wait-KeyPress` | 659 | Press-Enter 待機（AutoPilot/AutoConfirm スキップ） |
| `Wait-NetworkReady` | 673 | Test-Connection で host 到達待ち |

### A.5 権限・環境

| 関数 | 行番号 | 用途 |
|---|---|---|
| `Test-AdminPrivilege` | 3438 | 管理者権限判定 (bool) |

---

## B. 内部実装（PATCH で変更されうる、モジュール依存禁止）

### B.1 オーケストレーション

| 関数 | 場所 | 用途 |
|---|---|---|
| `Invoke-SafeCommand` | common 900 | 同期実行 + ModuleResult 捕捉 + 例外吸収 |
| `Invoke-SafeCommandAsync` | common 1008 | Runspace 実行 + Skip flag / Timeout 監視 |
| `Get-FabriqAsyncConfig` | common 984 | async_config.json 読み込み |
| `Invoke-BatchExecution` | main 223 | プロファイル一括実行ループ（マーカー処理含む） |
| `Invoke-KittingScript` | main 136 | 単発モジュール実行 |
| `Invoke-FlexProfileLoop` | main 561 | FlexProfile sub-loop（dashboard 駆動） |
| `Invoke-WindowsUpdateLoop` | main 872 | WU リブートループ（wu_state.json） |
| `Set-WindowsUpdateAutoLogon` | main 813 | WU 用 AutoLogon 設定 |
| `Clear-WindowsUpdateAutoLogon` | main 853 | WU 完了後の credential クリア |
| `Show-WindowsUpdateSummary` | main 1030 | WU 結果集計表示 |

### B.2 プロファイル解決

| 関数 | 場所 | 用途 |
|---|---|---|
| `Resolve-ProfileModules` | common 3105 | profile CSV → module list 変換（マーカー解釈含む） |
| `Initialize-ModuleSystem` | common 3887 | modules/{std,ext}/*/module.csv 自動検出 |
| `Build-CategoryMenu` | common 3866 | カテゴリ別グルーピング + 順序 |
| `Load-Profiles` | common 3063 | profiles/*.csv 一覧化 |
| `Create-DefaultProfiles` | common 3031 | Basic Setup / Full Setup の自動生成 |
| `Show-ProfileMenu` | common 3275 | コンソール用プロファイル一覧表示 |
| `Show-ProfileConfirmation` | common 3295 | プロファイル実行前確認 + AutoPilot 銘確認 |

### B.3 セッション・状態

| 関数 | 場所 | 用途 |
|---|---|---|
| `Initialize-Session` | common 2905 | worker / media serial / session.json |
| `Reset-FabriqState` | common 2781 | New Session / Refabriq の総リセット |
| `Get-VolumeSerial` | common 2892 | ドライブシリアル取得（media identification） |
| `Get-HardwareUniqueId` | common 739 | BIOS Serial > MAC > UNKNOWN の優先順位 |
| `Initialize-EvidenceBasePath` | common 785 | `{ts}_{pc}_{sn}_evidence` ディレクトリ命名 |

### B.4 Resume / Restart

| 関数 | 場所 | 用途 |
|---|---|---|
| `Save-ResumeState` | common 2681 | 環境変数 + Profile 状態 + DPAPI passphrase を json 保存 |
| `Load-ResumeState` | common 2766 | resume_state.json 読み込み |
| `Remove-ResumeState` | common 2775 | resume_state.json 削除 |
| `Restore-HostEnvironment` | common 2881 | json → env vars 復元 |
| `Register-FabriqRunOnce` | common 3970 | HKLM RunOnce に Fabriq.exe を登録 |
| `Register-FabriqActiveSetup` | common 3998 | Active Setup（GUID + StubPath）登録 |
| `Deploy-FabriqUserSetupLauncher` | common 4038 | C:\ProgramData\fabriq\fabriq_user_setup.ps1 配置 |
| `Deploy-FabriqStartupTrigger` | common 4114 | Default Profile Startup folder に .cmd 配置 |
| `Invoke-CountdownRestart` | common 4148 | 5 秒カウント → Restart-Computer -Force |
| `Invoke-AutoResumeCountdown` | common 4164 | AutoPilot 復帰確認（Enter/Y/Esc/N 制御） |
| `Invoke-CountdownSignout` | common 4349 | サインアウト用カウントダウン |
| `Wait-SystemReady` | common 692 | services 起動待ち + network 到達確認 |

### B.5 履歴・エビデンス

| 関数 | 場所 | 用途 |
|---|---|---|
| `Initialize-ExecutionHistory` | common 1422 | 履歴 CSV ディレクトリ生成 + 旧 path migration |
| `Write-ExecutionHistory` | common 1484 | 1 行追記 + retry 3 回 |
| `Import-ExecutionHistory` | common 1547 | KanriNo フィルタ + Limit + Timestamp 降順 |
| `Restore-ExecutionHistory` | common 1603 | 過去履歴を $script:ExecutionResults へ復元 |
| `Show-ExecutionHistory` | common 1713 | コンソール履歴表示 |
| `Export-ExecutionHistory` | common 1763 | logs/history + evidence/{base}/export_history への dual 出力 |
| `Export-HtmlChecklist` | common 1818 | HTML 監査レポート生成（Network / Printer 照合 + License + BitLocker + チェックリスト） |
| `Add-ExecutionResult` | common 1230 | $script:ExecutionResults に 1 件追加 + status.json 更新 |
| `Clear-ExecutionResults` | common 1257 | 結果配列クリア（IsRestored は保持） |
| `Show-ExecutionSummary` | common 1274 | 実行結果サマリ表示 |
| `Capture-ScreenEvidence` | common 4218 | 自動 PNG（auto_capture/、SetProcessDPIAware） |
| `Save-Screenshot` | common 4297 | 手動 PNG（gyotaku/、DPI 不変更） |
| `Complete-ProfileExecution` | common 2368 | finalize: Export-ExecutionHistory + Export-HtmlChecklist + log_uploader |

### B.6 ステータスモニタ

| 関数 | 場所 | 用途 |
|---|---|---|
| `Write-StatusFile` | common 3678 | status.json atomic write（PCInfo + CurrentPCInfo + Execution） |
| `Remove-StatusFile` | common 3763 | status.json + tmp + art_pulse.txt 削除 |
| `Start-StatusMonitor` | common 3782 | 別プロセスで kernel/ps1/status_monitor.ps1 起動 |
| `Stop-StatusMonitor` | common 3822 | CloseMainWindow → 2s wait → Kill |
| `Write-ArtPulse` | common 267 | art_pulse.txt の counter +1 |
| `Get-CurrentPCInfo` | common 3590 | Get-NetAdapter / Get-Printer から実 OS 状態取得 |

### B.7 暗号化補助

| 関数 | 場所 | 用途 |
|---|---|---|
| `Test-MasterPassphrase` | common 508 | passphrase_verify.txt を復号して "surkitinisme" と一致するか |
| `Protect-PassphraseForResume` | common 407 | DPAPI LocalMachine で passphrase を暗号化 → Base64 |
| `Unprotect-PassphraseFromResume` | common 424 | DPAPI で復号 |

### B.8 ユーザー権限・環境

| 関数 | 場所 | 用途 |
|---|---|---|
| `_Resolve-LoggedOnUser` | common 3458 | UAC elevation 時に `Win32_ComputerSystem.UserName` から logged-on user を解決 |
| `Expand-UserEnvironmentVariables` | common 3489 | %USERPROFILE% / %LOCALAPPDATA% / %APPDATA% を logged-on user 視点で展開 |
| `Resolve-HkcuRoot` | common 3554 | HKCU root の解決（admin elevation 環境用） |
| `Remove-ZoneIdentifier` | common 3523 | NTFS Zone.Identifier ADS 削除（Internet zone block 解除） |
| `Get-ModuleBasePath` | common 3430 | $PSScriptRoot or Get-Location |

### B.9 コンソール制御

| 関数 | 場所 | 用途 |
|---|---|---|
| `Hide-ConsoleWindow` | common 92 | ShowWindow(SW_HIDE) |
| `Show-ConsoleWindow` | common 99 | ShowWindow(SW_SHOW) + SetForegroundWindow |
| `Set-ConsoleForeground` | common 145 | foreground にだけする |
| `Set-ConsoleSize` | common 217 | Window 75x35, Buffer 75x9999 |
| `Disable-QuickEditMode` | common 127 | QuickEdit 無効化（クリックフリーズ防止） |
| `Enable-SleepSuppression` | common 238 | SetThreadExecutionState で sleep 抑止 |
| `Disable-SleepSuppression` | common 246 | sleep 抑止解除 |

### B.10 通知・エラー

| 関数 | 場所 | 用途 |
|---|---|---|
| `Invoke-ErrorNotification` | common 158 | 3-tone beep（Error）/ 2-tone beep（Partial） + foreground |
| `Show-AutoPilotErrorDialog` | common 189 | WinForms MessageBox で Retry/Cancel |

### B.11 ロギング

| 関数 | 場所 | 用途 |
|---|---|---|
| `Write-KitLog` | common 3385 | `[ts] [Level] Message` 形式コンソール出力（INFO/WARN/ERROR/SUCCESS） |
| `Save-RollbackInfo` | common 3404 | Category / Key / OldValue / NewValue を Write-KitLog 経由でロギング |
| `Clear-AllLogs` | common 2466 | logs ディレクトリ全削除 |
| `Clear-Evidence` | common 2589 | evidence ディレクトリ全削除 |

### B.12 CSV 補助

| 関数 | 場所 | 用途 |
|---|---|---|
| `Import-CsvSafe` | common 842 | Import-Csv -Encoding Default + retry + 0 行警告 |
| `Test-CsvColumns` | common 867 | 必須列の存在検証 |
| `Parse-MenuSelection` | common 1360 | "1,3-5;7" → @(1,3,4,5,7) のパース |
| `Test-BatchInput` | common 1389 | 入力が batch 形式か |
| `Show-BatchConfirmation` | common 1396 | batch 実行前確認 |
| `Show-Progress` | common 601 | `[N/Total] Activity (P%)` |
| `Show-BatchProgress` | common 611 | `=== N/Total : ItemName ===` |

### B.13 main.ps1 ローカル

| 関数 | 場所 | 用途 |
|---|---|---|
| `Load-HostList` | main 57 | hostlist.csv 読み込み（Default encoding） |
| `Set-SelectedHostEnvironment` | main 80 | hostlist 行 → SELECTED_* env vars（ENC 復号付き） |

### B.14 終了

| 関数 | 場所 | 用途 |
|---|---|---|
| `Exit-Fabriq` | common 3840 | Stop-StatusMonitor + Disable-SleepSuppression + Stop-Transcript（idempotent guard） |

---

## 関数数の総計

- **公開 API（モジュール依存可）**: 約 **18 関数**
- **内部実装（PATCH で変更されうる）**: 約 **75 関数**
- **合計**: 約 **93 関数**（common.ps1 + main.ps1）

公開 API は全体の **20% 弱**。残り 80% は内部実装で、kernel 開発者の自由度を確保する設計。

---

## なぜこの分担が機能するか

公開 API を絞ることで:

1. **モジュール側の依存ポイントが少なく、リファクタが容易**
2. **`KERNEL_API.md` の真のソース性が成立**（記載されていれば公開、されていなければ内部）
3. **PATCH バージョン昇格でモジュール再配備が不要**（公開 API 不変だから）
4. **MINOR 昇格時もモジュールは opt-in で新機能を使うか選べる**（既存モジュールは無風）

`Resolve-ProfileModules` のような公開しない関数は、引数追加・戻り値追加・内部 logic の刷新を自由に行える。実際 kernel 3.x 系では `IncludeDisabled` スイッチや FlexProfile 用の戻り値拡張が PATCH で行われた。
