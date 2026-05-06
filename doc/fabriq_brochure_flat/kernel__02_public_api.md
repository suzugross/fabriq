# カーネル公開 API（モジュールから利用可能なサーフェス）

`kernel/KERNEL_API.md` で公式宣言されているサーフェスの解説。fabriq モジュールが安全に依存できる関数・グローバル変数・環境変数・契約の全集合。

---

## §1 公開関数

### §1.1 表示・通知（color-coded console output）

| 関数 | 用途 | 副作用 |
|---|---|---|
| `Show-Info -Message <string>` | シアンで `[INFO] ...` を出力 | `Write-ArtPulse` で `art_pulse.txt` のカウンタを +1（モニタ画面の鼓動演出） |
| `Show-Success -Message <string>` | グリーンで `[SUCCESS] ...` | 同上 |
| `Show-Warning -Message <string>` | イエローで `[WARNING] ...` | 同上 |
| `Show-Error -Message <string>` | レッドで `[ERROR] ...` | 同上 |
| `Show-Skip -Message <string>` | ダークグレーで `[SKIP] ...` | 同上 |
| `Show-Separator` | シアンで横線 `========================================` | なし |
| `Show-CategorySeparator -Name <string>` | シアンで `=== <Name> ===` | なし |

**禁止事項**: モジュールは `Write-Host` を直接使ってはならない（CLAUDE.md ルール 2）。色付け・ART pulse・将来的な GUI ログ転送をすべて common 経由で得るため。

### §1.2 CSV 読み込み

```powershell
Import-ModuleCsv -Path <string> [-FilterEnabled] [-RequiredColumns <string[]>] [-Segment <string>]
```

**動作（4 ステップ統合パイプライン）**:

1. `Import-CsvSafe` で UTF-8/Default 自動判定、エラー時は `$null` 返却（呼び元で `Show-Skip` 系へ flow）
2. パスフレーズ（`$global:FabriqMasterPassphrase`）が立っていれば、各セルの `ENC:<Base64>` 値を `Unprotect-FabriqValue` で**透過復号**（モジュールは平文を受け取る）
3. `-RequiredColumns` 指定時は `Test-CsvColumns` で必須列を検証し、欠落時は `Show-Error` + `$null` 返却
4. `-FilterEnabled` 指定時は `Enabled -eq "1"` で絞り込み + `Segment` 列があれば `$env:FABRIQ_SEGMENT` で**厳密マッチ**フィルタ（空 vs 空もマッチ）

**Segment フィルタ仕様**: profile CSV の `Segment` 列の値が `$env:FABRIQ_SEGMENT` に渡され、`<name>_list.csv` の `Segment` 列と完全一致する行のみが返る。同モジュールを設定値別に呼び分けるため（例: `wallpaper_config` を `[seg:office]` と `[seg:home]` で別 CSV 行として実行）。

### §1.3 結果返却（モジュール契約の核）

```powershell
New-ModuleResult -Status <Success/Error/Cancelled/Skipped/Partial> -Message <string> [-Details <array>] [-Verified <Nullable[bool]>]
```

`PSCustomObject` を返却。フィールド: `_IsModuleResult=$true`（識別マーカー）, `Status`, `Message`, `Details`, `Verified`, `Timestamp`。同時に `$global:_LastModuleResult` にも stash（pipeline capture failure に対するフォールバック）。

```powershell
New-BatchResult -Success <int> -Skip <int> -Fail <int> [-Title <string>] [-MessageSuffix <string>] [-Verified <Nullable[bool]>]
```

集計を画面表示してから `Status` を自動判定（Fail=0 + Success>0 → Success / Success>0 + Fail>0 → Partial / Skip 全件 → Skipped / Fail のみ → Error）し、`New-ModuleResult` を内部で呼ぶ便利関数。

```powershell
Confirm-ModuleExecution [-Message <string>]
```

実行前 Y/N 確認。`N` で `Cancelled` の `ModuleResult` を返却（モジュール先頭で `if ($r = Confirm-ModuleExecution) { return $r }` パターン）。AutoPilot / AutoConfirm 中は自動 Y。

### §1.4 ユーザー確認・待機

| 関数 | 動作 | AutoPilot / AutoConfirm 挙動 |
|---|---|---|
| `Confirm-Execution [-Message]` | Y/N → `[bool]` 返却 | 自動 `$true`（コンソールに `[AUTOPILOT]` / `[AUTOCONFIRM]` を出力） |
| `Wait-KeyPress [-Message]` | Press-Enter 待ち | 何もせずに即 return（unattended 続行） |
| `Wait-NetworkReady [-Target] [-RetryIntervalSec] [-PingCount]` | `Test-Connection` でホスト到達待ち。デフォルト `8.8.8.8`、10 秒間隔、無限ループ（Ctrl+C で abort） | 同上の挙動（pingリトライは継続） |

### §1.5 権限・環境

| 関数 | 戻り値 | 用途 |
|---|---|---|
| `Test-AdminPrivilege` | `[bool]` | `WindowsPrincipal.IsInRole(Administrator)` を内部で評価 |
| `Unprotect-FabriqValue -EncryptedValue <string> -Passphrase <string>` | `[string]` | `ENC:` 値の AES-256-CBC 復号（`Import-ModuleCsv` で自動適用されるため通常モジュール側で直接呼ぶ必要はない） |

---

## §2 公開グローバル変数（読み取り専用）

| 変数 | 型 | 用途 |
|---|---|---|
| `$global:FabriqMasterPassphrase` | `string` | マスターパスフレーズ。`Unprotect-FabriqValue` 第二引数に渡す等の用途 |
| `$global:AutoPilotMode` | `bool` | プロファイル一括実行が AutoPilot か |
| `$global:AutoPilotWaitSec` | `int` | AutoPilot のモジュール間ウェイト秒（デフォ 3） |
| `$global:AutoConfirmMode` | `bool` | FlexProfile 単発実行（`[Run This]`）中か。AutoPilot のサブセット動作（Y/N と Press-Enter のみ短絡） |
| `$global:FabriqEvidenceBasePath` | `string` | エビデンス保存先ベース（`{Timestamp}_{PCName}_{Serial}_evidence/evidence/`）|

**書き込み可能なグローバルは `$global:_LastModuleResult` のみ**（`New-ModuleResult` が内部で更新するフォールバック）。

---

## §3 公開環境変数

### §3.1 選択ホスト情報（`Set-SelectedHostEnvironment` が hostlist.csv から流し込む）

```
SELECTED_KANRI_NO          管理 ID
SELECTED_OLD_PCNAME        旧 PC 名
SELECTED_NEW_PCNAME        新 PC 名（hostname_config 適用先）
SELECTED_ETH_IP            イーサネット IPv4
SELECTED_ETH_SUBNET        イーサネットサブネットマスク
SELECTED_ETH_GATEWAY       イーサネットゲートウェイ
SELECTED_WIFI_IP           Wi-Fi IPv4
SELECTED_WIFI_SUBNET       Wi-Fi サブネットマスク
SELECTED_WIFI_GATEWAY      Wi-Fi ゲートウェイ
SELECTED_DNS1..SELECTED_DNS4   DNS サーバ最大 4 件
SELECTED_PIN               セットアップ時の PIN（cert_config 等で参照）
SELECTED_PRINTER_<N>_NAME       N=1..10
SELECTED_PRINTER_<N>_DRIVER
SELECTED_PRINTER_<N>_PORT
```

`hostlist.csv` の `ENC:` フィールドはホスト選択時点で復号されるため、モジュール側はそのまま平文を読める。

### §3.2 プロファイル実行パラメータ

| 変数 | 由来 | 用途 |
|---|---|---|
| `FABRIQ_SEGMENT` | profile CSV の `Segment` 列 | `Import-ModuleCsv` の Segment フィルタに連動 |
| `FABRIQ_AUTOLOGON_USER` | `__AUTO_to_<User>__` マーカー | `autologon_config` モジュール専用、対象 User を指定 |
| `FABRIQ_WORKER_NAME` | session.json の `WorkerName` | 履歴/エビデンスメタデータ |
| `FABRIQ_EVIDENCE_BASE` | `Initialize-EvidenceBasePath` | `$global:FabriqEvidenceBasePath` と同値（モジュール内部で `Join-Path` 用に使う） |

---

## §4 Profile CSV スキーマ（モジュール開発者向け契約）

### §4.1 列定義

| 列 | 必須 | 用途 |
|---|---|---|
| `Order` | 必須 | 整数・昇順実行。実行履歴の一級識別子（同一 MenuName が複数行ある時の区別に使用） |
| `ScriptPath` | 必須 | `{standard,extended}/<module>/<script>.ps1` 形式 or 特殊マーカー。区切りは `/` `\` どちらも可 |
| `Enabled` | 必須 | `1`=実行 / `0`=スキップ |
| `Description` | 任意 | プロファイル UI 表示用コメント。`__AUTOPILOT__` 行では `WaitSec=N` 形式で wait 秒指定 |
| `Segment` | 任意 | `Import-ModuleCsv` の Segment フィルタ値として渡される |
| `ErrorMode` | 任意 | AutoPilot 時のエラー処理（空=ダイアログ確認 / `skip` / `retry` 最大 5 回） |
| `Group` | 任意（kernel 3.2.0〜） | FlexProfile dashboard の Groups バー集約名。Linear `Execute Profile` は無視 |

### §4.2 特殊マーカー（5 種、kernel 3.0.0 で 4 種削除済み）

| マーカー | 動作 | 導入版 |
|---|---|---|
| `__AUTOPILOT__` | 以降を AutoPilot 化（Y/N 自動承認 + 指定 wait 秒のモジュール間スリープ） | 2.0.0 |
| `__ASYNC__` | 以降を Runspace 実行に切り替え。Status Monitor の Skip ボタン or `async_config.json` の `DefaultTimeoutSec` で強制中断可能 | 2.1.0 |
| `__RESTART__` | Windows 再起動 → RunOnce 経由で resume | 2.0.0 |
| `__REEXPLORER__` | Explorer 再起動（HKCU レジストリ変更の即時反映等） | 2.0.0 |
| `__AUTO_to_<User>__` | `autologon_config` を該当 User で呼び出し | 2.0.0 |

**3.0.0 で削除**（破壊的変更 / MAJOR 昇格）: `__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__`。旧プロファイルが含んでいても `Resolve-ProfileModules` の `$invalidPaths` 経由で「module not found」warning に降格、他モジュールは続行（graceful degradation）。

### §4.3 Group 列セマンティクス（kernel 3.2.0〜）

- 同一 `Group` 値の行群を FlexProfile dashboard の `[Run: <Group>]` ボタンで一括実行
- 実行は AutoPilot 挙動 + `FinalizeOnComplete:$false`（完了は operator が `[Complete]` で手動）
- Group 跨ぎ間の `__RESTART__` は当該 Group 実行時には skip（**literal interpretation**: Group 列が batch を厳密に決定）

---

## §5 ModuleResult 契約

モジュールは実行後、**pipeline 経由で** `New-ModuleResult` または `New-BatchResult` を返却する。

### フィールド

| フィールド | 型 | 意味 |
|---|---|---|
| `_IsModuleResult` | `bool` | 識別マーカー（常に `$true`） |
| `Status` | `string` | `Success` / `Error` / `Cancelled` / `Skipped` / `Partial` |
| `Message` | `string` | 結果メッセージ |
| `Details` | `array` | 任意の詳細情報（実行履歴 CSV には入らない） |
| `Verified` | `Nullable[bool]` | Post-Apply Verification 結果。`$null` = 未検証、`$true` = PASS、`$false` = FAIL |
| `Timestamp` | `DateTime` | 生成日時 |

### 標準呼び出しパターン

```powershell
return (New-ModuleResult -Status "Success" -Message "Done" -Verified $true)
return (New-BatchResult -Success 3 -Skip 1 -Fail 0 -Title "Foo Results" -Verified $true)
```

カーネル側 `Invoke-SafeCommand` / `Invoke-KittingScript` は pipeline output を走査して `_IsModuleResult -eq $true` の要素を捕捉する。pipeline capture が失敗した場合は `$global:_LastModuleResult` からフォールバック取得する二重防御。

### Status セマンティクス（実行履歴 CSV / HTML チェックリスト / Status Monitor 共通）

| Status | 意味 | HTML 上の色 |
|---|---|---|
| `Success` | 正常完了 | 緑 |
| `Partial` | 一部成功・一部失敗（`New-BatchResult` は Success>0 + Fail>0 で自動付与） | 黄 |
| `Skipped` | 全件スキップ（対象なし、または明示的に skip） | 灰 |
| `Cancelled` | ユーザーキャンセル（Y/N で N 押下） | 黄 |
| `Error` | エラー発生 | 赤 |

### Verified セマンティクス（Post-Apply Verification）

- `$null`: 検証未実装または検証不可（例: `sysprep_config` は再起動後にしか検証できない）
- `$true`: 適用後に読み返した実 OS 状態が期待値と完全一致（PASS）
- `$false`: 適用は受理されたが再読み込み時に乖離（FAIL）

検証が技術的に困難 or 偽 PASS の risk があるモジュールは `-Verified` を省略する（`$null`）。fabriq では `acl_config` / `spi_config` / `copyfile_config` 等が偽 PASS リスクで除外済み。
