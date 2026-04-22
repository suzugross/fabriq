# Fabriq Kernel Public API

**Current Kernel Version**: `2.0.0`（`kernel/KERNEL_VERSION` を真のソースとする）

このドキュメントで「公開 API」として宣言されている要素のみが、モジュールから依存してよいカーネル機能です。ここに記載されていない `common.ps1` 関数・グローバル変数・内部状態ファイルは**内部実装**であり、PATCH バージョンでも予告なく変更される可能性があります。

公開 API の追加・削除・シグネチャ変更は **必ず `KERNEL_VERSION` の昇格（MINOR 以上）と同コミット内で** 本ドキュメントに反映します（CLAUDE.md ルール G）。

---

## 1. 公開関数（モジュールから呼び出し可）

### 1.1 表示・通知
| 関数 | シグネチャ | 用途 |
|---|---|---|
| `Show-Info` | `-Message <string>` | 情報表示（シアン） |
| `Show-Success` | `-Message <string>` | 成功表示（グリーン） |
| `Show-Warning` | `-Message <string>` | 警告表示（イエロー） |
| `Show-Error` | `-Message <string>` | エラー表示（レッド） |
| `Show-Skip` | `-Message <string>` | スキップ表示（グレー） |
| `Show-Separator` | なし | 区切り線表示 |
| `Show-CategorySeparator` | `-Name <string>` | カテゴリ区切り表示 |

### 1.2 CSV 読み込み
| 関数 | シグネチャ | 用途 |
|---|---|---|
| `Import-ModuleCsv` | `-Path <string> [-FilterEnabled] [-RequiredColumns <string[]>] [-Segment <string>]` | モジュール用 CSV を透過復号・列検証・Segment フィルタ付きで読み込む |

**契約**: `Segment` 列を持つ CSV は `$env:FABRIQ_SEGMENT` で厳密一致フィルタされる（空 vs 空もマッチ）。

### 1.3 結果返却
| 関数 | シグネチャ | 用途 |
|---|---|---|
| `New-ModuleResult` | `-Status <Success/Error/Cancelled/Skipped/Partial> -Message <string> [-Details <array>] [-Verified <Nullable[bool]>]` | モジュール結果オブジェクト生成 |
| `New-BatchResult` | `-Success <int> -Skip <int> -Fail <int> [-Title <string>] [-MessageSuffix <string>] [-Verified <Nullable[bool]>]` | 集計表示 + Status 自動判定で `New-ModuleResult` を返却 |
| `Confirm-ModuleExecution` | `[-Message <string>]` | 実行前 Y/N 確認。`N` で `Cancelled` の ModuleResult を返却。AutoPilot 中は自動 Y |

### 1.4 ユーザー確認・待機
| 関数 | シグネチャ | 用途 |
|---|---|---|
| `Confirm-Execution` | `[-Message <string>]` | Y/N 確認（bool 返却）。AutoPilot 中は自動 Y |
| `Wait-KeyPress` | `[-Message <string>]` | キー入力待ち。AutoPilot 中はスキップ |
| `Wait-NetworkReady` | `[-Target <string>] [-RetryIntervalSec <int>] [-PingCount <int>]` | ネットワーク到達待ち |

### 1.5 権限・環境
| 関数 | シグネチャ | 用途 |
|---|---|---|
| `Test-AdminPrivilege` | なし（bool 返却） | 管理者権限判定 |
| `Unprotect-FabriqValue` | `-EncryptedValue <string> -Passphrase <string>` | `ENC:` 値の復号（`Import-ModuleCsv` で自動適用されるため通常不要） |

---

## 2. 公開グローバル変数（モジュールから読み取り可）

| 変数 | 型 | 用途 |
|---|---|---|
| `$global:FabriqMasterPassphrase` | `string` | マスターパスフレーズ（`Unprotect-FabriqValue` に渡す際などに使用） |
| `$global:AutoPilotMode` | `bool` | AutoPilot 実行中か |
| `$global:AutoPilotWaitSec` | `int` | AutoPilot モジュール間ウェイト秒 |
| `$global:FabriqEvidenceBasePath` | `string` | エビデンス保存先ベースパス（サブディレクトリを作る基点） |

**書き込み**: 公開 API として書き込み可能なグローバル変数は `$global:_LastModuleResult`（`New-ModuleResult` 内部で自動更新）のみ。それ以外は読み取り専用。

---

## 3. 公開環境変数（モジュールから参照可）

### 3.1 選択ホスト情報（Set-SelectedHostEnvironment が設定）
- `SELECTED_KANRI_NO` / `SELECTED_OLD_PCNAME` / `SELECTED_NEW_PCNAME`
- `SELECTED_ETH_IP` / `SELECTED_ETH_SUBNET` / `SELECTED_ETH_GATEWAY`
- `SELECTED_WIFI_IP` / `SELECTED_WIFI_SUBNET` / `SELECTED_WIFI_GATEWAY`
- `SELECTED_DNS1` / `SELECTED_DNS2` / `SELECTED_DNS3` / `SELECTED_DNS4`
- `SELECTED_PIN`
- `SELECTED_PRINTER_<N>_NAME` / `SELECTED_PRINTER_<N>_DRIVER` / `SELECTED_PRINTER_<N>_PORT`（N=1..10）

### 3.2 プロファイル実行パラメータ
- `FABRIQ_SEGMENT` — `Import-ModuleCsv` の Segment フィルタ対象値
- `FABRIQ_AUTOLOGON_USER` — `__AUTO_to_<User>__` マーカーで渡される User 名（autologon_config 専用）
- `FABRIQ_WORKER_NAME` — 選択中の作業者名
- `FABRIQ_EVIDENCE_BASE` — エビデンスベースパス（`$global:FabriqEvidenceBasePath` と同値）

---

## 4. Profile CSV スキーマ（モジュール開発者向け契約）

### 4.1 必須・任意列
| 列 | 必須 | 用途 |
|---|---|---|
| `Order` | 必須 | 整数・昇順実行 |
| `ScriptPath` | 必須 | `{standard,extended}/<module>/<script>.ps1` 形式 or 特殊マーカー |
| `Enabled` | 必須 | `1` = 実行 / `0` = スキップ |
| `Description` | 任意 | 表示・メモ |
| `Segment` | 任意 | `Import-ModuleCsv` の Segment フィルタ値として渡される |
| `ErrorMode` | 任意 | AutoPilot 時のエラー処理（空 / `skip` / `retry`） |

### 4.2 特殊マーカー
| マーカー | 動作 |
|---|---|
| `__AUTOPILOT__` | 以降を AutoPilot 化（`Description` に `WaitSec=N`） |
| `__RESTART__` | Windows 再起動 + RunOnce 経由で再開 |
| `__SHUTDOWN__` | Windows シャットダウン |
| `__PAUSE__` | ユーザー入力待ち |
| `__REEXPLORER__` | Explorer 再起動 |
| `__STOPLOG__` / `__STARTLOG__` | Transcript 停止・再開 |
| `__AUTO_to_<User>__` | `autologon_config` に User 指定で呼び出し |

---

## 5. ModuleResult 契約

モジュールは実行後、**pipeline 経由で** `New-ModuleResult` または `New-BatchResult` を返却する。

### 5.1 フィールド
| フィールド | 型 | 意味 |
|---|---|---|
| `_IsModuleResult` | `bool` | 識別マーカー（常に `$true`） |
| `Status` | `string` | `Success` / `Error` / `Cancelled` / `Skipped` / `Partial` |
| `Message` | `string` | 結果メッセージ |
| `Details` | `array` | 任意の詳細情報 |
| `Verified` | `Nullable[bool]` | Post-Apply Verification 結果（未実施は `$null`） |
| `Timestamp` | `DateTime` | 生成日時 |

### 5.2 呼び出しパターン
```powershell
return (New-ModuleResult -Status "Success" -Message "Done" -Verified $true)
return (New-BatchResult -Success 3 -Skip 1 -Fail 0 -Title "Foo Results" -Verified $true)
```

---

## 6. 内部実装（非 API、PATCH で変更される可能性あり）

以下はカーネル内部実装であり、モジュールから依存してはいけません。変更しても `KERNEL_VERSION` は PATCH 昇格のみです。

- `Invoke-SafeCommand` / `Invoke-SafeCommandAsync` / `Invoke-BatchExecution` / `Invoke-KittingScript`
- `Resolve-ProfileModules` / `Initialize-ModuleSystem` / `Build-CategoryMenu` / `Load-Profiles`
- `Save-ResumeState` / `Load-ResumeState` / `Remove-ResumeState` / `Restore-HostEnvironment`
- `Register-FabriqRunOnce` / `Invoke-CountdownRestart` / `Invoke-CountdownShutdown`
- `Write-ExecutionHistory` / `Initialize-ExecutionHistory` / `Restore-ExecutionHistory` / `Export-ExecutionHistory` / `Export-HtmlChecklist`
- `Capture-ScreenEvidence` / `Initialize-EvidenceBasePath`
- `Write-StatusFile` / `Start-StatusMonitor` / `Stop-StatusMonitor`
- `Protect-PassphraseForResume` / `Unprotect-PassphraseFromResume`
- `Test-MasterPassphrase` / `Add-ExecutionResult` / `Clear-ExecutionResults` / `Show-ExecutionSummary`
- 状態ファイル: `kernel/json/resume_state.json`, `status.json`, `session.json`, `art_pulse.txt`, `async_config.json`, `skip_request.flag`
- オーケストレータ経由で設定される仕組み（`__ASYNC__` の Runspace 実装等）
- Status Monitor の UI 構成・ボタン配置
- HTML チェックリストのテンプレート

---

## 7. 変更時の運用

| 変更種別 | KERNEL_VERSION 影響 | KERNEL_API.md 更新 | 既存モジュール影響 |
|---|---|---|---|
| 公開 API の削除・シグネチャ破壊変更・フィールド削除 | **MAJOR** | 必須 | **全件監査必須** |
| 公開 API の追加（関数・マーカー・任意引数・列・環境変数） | **MINOR** | 必須 | なし（opt-in） |
| 内部実装のみの修正 | **PATCH** | 不要 | なし |
| ドキュメント・コメントのみ | 変更なし | 不要 | なし |

詳細運用は `CLAUDE.md` の「バージョン管理ルール」セクションを参照。
