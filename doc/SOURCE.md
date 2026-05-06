# Fabriq Brochure Materials — SOURCE

本ファイルは fabriq 技術パンフレット用素材を全 100 章を 1 ファイルに連結したもの。
ナビゲーションは別ファイル `INDEX.md` を参照してください。各章は `# === <path> ===` の見出しで開始します。

---


<!-- ============================================================ -->
# === kernel/01_overview.md ===
<!-- ============================================================ -->

# カーネル全体像

**現行版**: `kernel/KERNEL_VERSION` = `3.2.2`（fabriq ver3.2 — *Manifeste du Surkitinisme*）

## カーネルとは何か

`kernel/` は fabriq フレームワークの**運転系**である。モジュールを動かす実行基盤・GUI・状態管理・暗号化・ロギング・エビデンス取得・再起動跨ぎ・更新オーバーレイ契約のすべてを抱え込む層であり、モジュール側からは「公開 API」（`KERNEL_API.md`）で宣言されたサーフェスのみが利用可能。

**役割の分担**:

```
Fabriq.exe (C#ランチャ)
   ↓ UAC 自動昇格
   ↓ 管理者権限の PowerShell コンソールを開く
kernel/main.ps1
   ↓ . source
kernel/common.ps1（90+ 関数の共通ライブラリ）
   ↓ dot-source
apps/fabriq_operator/fabriq_operator.ps1（WinForms ダッシュボード）
   ↓ 操作
modules/{standard,extended}/<name>/<name>.ps1（実モジュール）
```

## カーネルが担う責務（カテゴリ別）

| カテゴリ | 責務 | 主な関数・ファイル |
|---|---|---|
| **エントリポイント** | UAC 昇格 → 管理者権限取得 → main.ps1 起動 → GUI ダッシュボードへ受け渡し | `Fabriq.exe`（C#）, `kernel/main.ps1` |
| **共通ライブラリ** | 表示・確認ダイアログ・CSV 読み込み・結果オブジェクト生成・暗号化・履歴・エビデンス取得・スリープ抑制・コンソール制御の共通関数群（90+） | `kernel/common.ps1` |
| **モジュールシステム** | `modules/{standard,extended}/<name>/module.csv` の自動検出・カテゴリ別グルーピング・カテゴリ順序管理 | `Initialize-ModuleSystem`, `Build-CategoryMenu`, `kernel/csv/categories.csv` |
| **プロファイル解決** | profile CSV の `Order` / `ScriptPath` / `Enabled` / `Segment` / `ErrorMode` / `Group` を読み込み、特殊マーカーを解釈してモジュールリストへ展開 | `Resolve-ProfileModules`, `Load-Profiles` |
| **オーケストレーション** | プロファイル一括実行・モジュール単発実行・AutoPilot ループ・FlexProfile sub-loop・Windows Update リブートループ | `Invoke-BatchExecution`, `Invoke-KittingScript`, `Invoke-FlexProfileLoop`, `Invoke-WindowsUpdateLoop` |
| **エラー制御** | モジュール例外を吸収して `ModuleResult` を取り出し、AutoPilot 中の `ErrorMode` (skip/retry/ask) を分岐 | `Invoke-SafeCommand`, `Invoke-SafeCommandAsync`, `Show-AutoPilotErrorDialog` |
| **CSV 読み込み + 暗号化透過復号** | `Import-CsvSafe` → 必須列検証 → `ENC:` 値の AES-256-CBC 復号 → `Segment` フィルタ → `Enabled` フィルタ | `Import-ModuleCsv`, `Unprotect-FabriqValue`, `Test-MasterPassphrase` |
| **再起動跨ぎ** | `__RESTART__` 検出 → 状態保存（`resume_state.json`）→ RunOnce 登録 → 再起動 → 復帰時に `Wait-SystemReady` → `Invoke-AutoResumeCountdown` → 環境復元 → 残モジュール継続 | `Save-ResumeState`, `Load-ResumeState`, `Register-FabriqRunOnce`, `Invoke-CountdownRestart`, `Invoke-AutoResumeCountdown`, `Restore-HostEnvironment` |
| **エビデンス自動収集** | モジュール実行ごとにスクリーンショット PNG 保存 + 実行履歴 CSV 追記 + プロファイル完了時に HTML チェックリスト生成 | `Capture-ScreenEvidence`, `Save-Screenshot`, `Write-ExecutionHistory`, `Export-HtmlChecklist`, `Initialize-EvidenceBasePath` |
| **ステータスモニタ** | 別プロセス（`status_monitor.ps1`）を起動し、`status.json` をリアルタイム書き込み → 別ウィンドウで進捗・PC 情報・ART pulse を可視化 | `Write-StatusFile`, `Start-StatusMonitor`, `Stop-StatusMonitor`, `Write-ArtPulse` |
| **AutoPilot / AutoConfirm** | プロファイル一括実行で Y/N 確認を自動承認、モジュール間ウェイトを設定 / FlexProfile 単発実行では Y/N と Press-Enter のみ短絡 | `$global:AutoPilotMode`, `$global:AutoConfirmMode`, `Confirm-Execution`, `Wait-KeyPress` |
| **セッション管理** | 作業者選択・媒体シリアル取得・`session.json` 保存・パスフレーズ検証 | `Initialize-Session`, `Test-MasterPassphrase`, `Reset-FabriqState` |
| **公開契約** | カーネル公開 API（§1〜§5）/ 更新オーバーレイ契約（§9）/ Evidence Manifest 契約（§10） | `KERNEL_API.md`, `dev/framework_overlay_rules.json`, `kernel/EVIDENCE_MANIFEST.md` |

## 保管場所マップ（kernel 配下）

```
kernel/
├── KERNEL_VERSION          ── 3.2.2（カーネル API SemVer の真のソース）
├── KERNEL_API.md           ── 公開 API サーフェスの明文化（§1〜§11）
├── EVIDENCE_MANIFEST.md    ── manifest.json 公開契約（外部 evidence consumer 向け）
├── common.ps1              ── 90+ 関数の共通ライブラリ（4371 行）
├── main.ps1                ── エントリスクリプト・FlexProfile sub-loop・Windows Update ループ（1913 行）
├── csv/
│   ├── categories.csv      ── カテゴリと表示順マスタ
│   ├── hostlist.csv        ── 対象 PC マスタ（暗号化フィールド対応）
│   ├── workers.csv         ── 作業者マスタ
│   ├── log_destinations.csv── ログ配送先マスタ（log_uploader 用）
│   └── manifesto.csv       ── マニフェスト本文（演出機能）
├── json/
│   ├── status.json         ── ステータスモニタ用ライブ状態（atomic write）
│   ├── session.json        ── 現セッション情報（worker, media serial, start time）
│   ├── resume_state.json   ── 再起動跨ぎ時の状態スナップショット（v1/v2 schema）
│   ├── async_config.json   ── __ASYNC__ Runspace 制御パラメータ
│   ├── art_pulse.txt       ── 動作鼓動カウンタ（演出用、Show-* で +1）
│   └── skip_request.flag   ── async モジュール強制スキップ要求の flag ファイル
├── ps1/
│   ├── status_monitor.ps1  ── 別プロセス WinForms モニタ（status.json を polling）
│   ├── view_report.ps1     ── HTML チェックリストの単体ビューア
│   ├── manifesto.ps1       ── マニフェスト表示 GUI
│   └── art_display.ps1     ── ART 演出（status_monitor に統合済）
└── txt/
    ├── passphrase_verify.txt ── パスフレーズ検証トークン（Studio で生成、起動必須）
    ├── art_sentences.txt   ── ART pulse で表示する一文集
    └── silence.flag        ── 演出抑制 flag（存在すれば ART を黙らせる）
```

## カーネル API のポリシー

- **真のソースは `kernel/KERNEL_VERSION`**。`README.md` L1 / `common.ps1` L2 / `main.ps1` L3 の版表記は `X.Y` 桁で同期。
- **公開 API は `KERNEL_API.md` の §1〜§5 のみ**。これに記載のない `common.ps1` 関数は内部実装で、PATCH バージョンでも予告なく変更されうる。
- **モジュール側は `REQUIRES_KERNEL` ファイル（1 行 `X.Y.Z`）で要求カーネル版を宣言**。更新オーバーレイ時に `REQUIRES_KERNEL > 現行 KERNEL_VERSION` ならカーネル先行更新を強制。
- **API 変更は `KERNEL_API.md` の同コミット更新が必須**（CLAUDE.md ルール G）。MINOR/MAJOR 昇格に必ず追従。


<!-- ============================================================ -->
# === kernel/02_public_api.md ===
<!-- ============================================================ -->

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


<!-- ============================================================ -->
# === kernel/03_orchestration.md ===
<!-- ============================================================ -->

# オーケストレーション層

カーネルがモジュールをどう順序付けて呼び出し、結果を集約するか。`main.ps1` のメインループと `common.ps1` の `Invoke-BatchExecution` / `Invoke-FlexProfileLoop` がコア。

---

## main.ps1 起動フロー

```
Fabriq.exe（C# ランチャ、UAC 自動昇格）
   ↓
kernel/main.ps1
   ├── . kernel/common.ps1            ── 共通関数読み込み
   ├── . kernel/ps1/manifesto.ps1     ── マニフェスト GUI 関数
   ├── . apps/fabriq_operator/fabriq_operator.ps1
   │                                  ── ダッシュボード GUI
   ├── Enable-SleepSuppression        ── SetThreadExecutionState で sleep 抑止
   ├── Disable-QuickEditMode          ── マウスクリックでフリーズしないよう QuickEdit 無効化
   ├── Set-ConsoleSize -Columns 75 -Lines 35
   ├── Start-Transcript               ── logs/{timestamp}_{uid}_{hostname}.log
   │
   ├── ── Resume Detection ──
   ├── Test-Path wu_state.json → Invoke-WindowsUpdateLoop（WU resume 専用、最優先）
   ├── Load-ResumeState → resume_state 読み込み
   │   ├── schemaVersion>=2 + ExecutionMode='Flex' → FlexProfile resume ルート
   │   ├── AutoPilot=true → Wait-SystemReady + Invoke-AutoResumeCountdown 60s
   │   └── 通常 → Confirm-Execution "Resume profile execution?"
   │   shouldResume なら:
   │     ├── Restore-HostEnvironment   ── SELECTED_* 環境変数を json から戻す
   │     ├── DPAPI 復号で master passphrase 復元 or 手動再入力（最大 3 回）
   │     └── EvidenceBasePath / SessionID 復元
   │
   ├── ── Fresh Start（resume なしの場合）──
   ├── Test-Path passphrase_verify.txt （無いと exit 1）
   ├── Load-HostList                  ── kernel/csv/hostlist.csv
   ├── Show-SessionSetupForm          ── WinForms で worker / host / passphrase を一括選択
   │   （3 つを 1 ダイアログで取る。GUI モード必須、CLI モードは廃止）
   ├── Set-SelectedHostEnvironment    ── SELECTED_* env vars を立てる
   ├── Initialize-EvidenceBasePath    ── evidence/{ts}_{pc}_{sn}_evidence/ を作る
   ├── Initialize-Session             ── workers / media serial / session.json
   ├── Initialize-ExecutionHistory    ── logs/history/execution_history.csv 作成 + 旧パスからの migrate
   ├── Initialize-ModuleSystem        ── modules/{standard,extended}/*/module.csv を再帰検出
   ├── Load-Profiles                  ── profiles/*.csv を一覧化（無ければ Basic Setup / Full Setup を生成）
   ├── Start-StatusMonitor            ── 別プロセスで status_monitor.ps1 を起動（PID 記録）
   │
   └── Show-Operator-Dashboard        ── apps/fabriq_operator のメインダッシュボード（無限ループ）
```

メインダッシュボードがイベントを受けるたびに、対応するアクション（個別モジュール実行・プロファイル実行・FlexProfile 起動・履歴表示・apps 起動・新セッション・終了）を呼び分ける。

---

## Invoke-BatchExecution（プロファイル一括実行の中核）

`common.ps1` 内、約 320 行のメイン関数。Linear `Execute Profile` も FlexProfile の `RunBatch` / `RunSingle` / `RunGroup` も最終的にここを通る。

### パラメータ

```powershell
Invoke-BatchExecution
    -SelectedModules <array>           # Resolve-ProfileModules が返した module オブジェクト配列
    [-AutoPilot]                       # AutoPilot モードか
    [-AutoPilotWaitSec <int>]          # モジュール間スリープ秒（デフォ 3）
    [-ProfilePath <string>]            # CSV パス（restart 時の resume 用）
    [-ProfileName <string>]            # 表示名
    [-FullProfileModules <array>]      # チェックリスト用フル module list
    [-ProfileStartTime <datetime>]     # 経過時間計算の絶対起点
    [-FinalizeOnComplete <bool>]       # Linear=true（HTML 自動生成）/ Flex=false（手動）
    [-ExecutionMode <Linear|Flex>]     # resume_state.json の schemaVersion を切替
    [-SelectedOrders <int[]>]          # FlexProfile が ResumeState に書く Order 配列
    [-ModuleStates <hashtable>]        # FlexProfile が ResumeState に書く状態マップ
```

### モジュール 1 件の処理ループ

```
foreach $module in $SelectedModules:
   1. progress 表示（Show-BatchProgress）
   2. 特殊マーカー分岐:
      ├── _IsRestart        → Save-ResumeState + Register-FabriqRunOnce + Invoke-CountdownRestart（return）
      └── _IsReexplorer     → Stop-Process explorer + 復活待ち → Add-ExecutionResult
   3. AutoPilot inter-module wait（current > 1 のみ）
   4. retry loop（do/while $retryModule）:
      a. _AutoLogonUser があれば $env:FABRIQ_AUTOLOGON_USER を立てる
      b. _Segment があれば $env:FABRIQ_SEGMENT を立てる
      c. _IsAsync && async_config.Enabled なら Invoke-SafeCommandAsync、
         さもなくば Invoke-SafeCommand
      d. 上記 env を解除
      e. result.Status が Error/Partial の場合:
         - Invoke-ErrorNotification（3-tone beep + console foreground）
         - AutoPilot 中なら _ErrorMode で分岐:
           ├── "skip"  → 記録して続行
           ├── "retry" → autoRetryCount++（最大 5 回）→ retryModule=true で再実行
           └── ""/"ask" → Show-AutoPilotErrorDialog（Retry / Skip）
   5. Add-ExecutionResult + Write-ExecutionHistory + Capture-ScreenEvidence
   6. $completedResults に Order/MenuName/Status/Verified/Message を追加
end foreach

7. Remove-ResumeState（再起動なしで完走できた場合の resume クリア）
8. Show-ExecutionSummary
9. Linear かつ ProfileName が指定されていれば:
   Complete-ProfileExecution（HTML チェックリスト生成 + log_uploader 起動 + view_report）
```

`finally` 節で必ず `$global:AutoPilotMode = $false` にリセット（Profile スコープ保証）+ `$script:LastBatchResults = $completedResults` を publish（FlexProfile dashboard が polling して状態反映）。

### ErrorMode マトリクス

| `_ErrorMode` 値 | AutoPilot=on | AutoPilot=off |
|---|---|---|
| 空文字 / `ask` | `Show-AutoPilotErrorDialog`（Retry/Skip 選択） | エラー記録のみ、続行 |
| `skip` | warning 出力 → エラー記録のまま続行 | 同左（off でも適用） |
| `retry` | autoRetryCount を最大 5 回まで増やして再実行、超えたらエラー記録 | 同左 |

---

## Invoke-FlexProfileLoop（FlexProfile sub-loop）

`main.ps1` 内、L561〜。FlexProfile dashboard と Invoke-BatchExecution の橋渡し。

### イベント駆動アクション一覧

dashboard が intent を返すたびに loop が dispatch する：

| Action | 動作 |
|---|---|
| `RunSingle` | 1 行だけ `Invoke-BatchExecution` で実行（`AutoConfirmMode=$true`、Y/N と Press-Enter のみ短絡） |
| `RunGroup` | `Group` 列で集約された Order 群を `AutoPilot=$true` + `FinalizeOnComplete:$false` で実行 |
| `RunBatch` | チェックボックスで選択された Order 群を `AutoPilot=$true` + `FinalizeOnComplete:$false` で実行 |
| `Complete` | `Complete-ProfileExecution -Mode 'Manual'`（HTML 生成 + log_uploader 発火） |
| `RestartNow` | プロファイル外 RESTART。`ResumeAfterOrder=-1` sentinel で resume_state を書き、`Register-FabriqRunOnce` + `Invoke-CountdownRestart` |
| `ResetState` | 該当 Order の状態を `Pending` に戻す（再実行候補） |
| `Close` | `Remove-ResumeState`（防御的）+ loop 脱出 |

### PendingFinalize フラグ

`RunBatch` / `RunSingle` / `RunGroup` / `ResetState` が走ったら `$pendingFinalize = $true`。dashboard は赤 PENDING FINALIZE バッジを点灯し、`Back` / `X` 押下時に確認ダイアログで警告。`Complete` で `$false` にリセット。

### 実行モデルの統一（kernel 3.1.5〜）

> **実行 = 常に AutoPilot 挙動 / 完了 = 常に手動**

AutoPilot トグルは UI から消した。operator の意思決定は「どのモジュールを動かすか」と「いつ Complete するか」の 2 軸のみ。これにより成果物確認 → 明示 finalize の運用に最適化されている。

---

## Resolve-ProfileModules（profile CSV → モジュールリスト変換）

`Invoke-BatchExecution` に渡す前段。約 170 行。

### 入力 → 出力

```
profile CSV (Order, ScriptPath, Enabled, Description, Segment, ErrorMode, Group)
   ↓
[ValidModules: array of module objects with attached _IsAsync / _Segment / _ErrorMode / _Group / _AutoLogonUser / _IsRestart / _IsReexplorer / _IsCheckedDefault / Order]
[InvalidPaths: array of unresolved ScriptPath strings]
[AutoPilot: bool, set if __AUTOPILOT__ row was Enabled=1]
[AutoPilotWaitSec: int, parsed from "WaitSec=N" in __AUTOPILOT__'s Description]
```

### 特殊マーカー処理

- `__AUTOPILOT__`: ValidModules には入れず、戻り値の `AutoPilot` フラグだけ立てる
- `__ASYNC__`: 同様に list には入れず、以降のモジュールに `_IsAsync=$true` を打つ sticky フラグ（プロファイル末尾まで継続）
- `__AUTO_to_<User>__`: `autologon_config` モジュールを参照し、`_AutoLogonUser` を attach、MenuName を `[AUTO:User] AutoLogon Configuration` に書き換え
- `__RESTART__` / `__REEXPLORER__`: 専用 PSCustomObject を生成し、`_IsRestart` / `_IsReexplorer` を立てる
- 通常 ScriptPath: AllModules リストから `RelativePath` 完全一致で探索、見つかれば copy + Order/Segment/ErrorMode/Group を attach

### IncludeDisabled スイッチ（FlexProfile 専用）

- 通常: `Enabled=1` 行のみ返す
- `-IncludeDisabled`: 全行返し、各 module に `_IsCheckedDefault`（Enabled の元値）を attach。dashboard はこれを見て初期チェックボックス状態を決定

`__AUTOPILOT__` / `__ASYNC__` マーカーは `IncludeDisabled` でも `Enabled=1` のときのみ効果を発揮する（disabled marker 行が global state を勝手に flip しない安全弁）。

---

## Invoke-WindowsUpdateLoop（独立した再起動ループ）

`main.ps1` L872。Windows Update のリブート跨ぎを `__RESTART__` と同じ仕組みで自動化。

### 制御の輪

```
windows_update.ps1（1 pass で WU を 1 サイクル回し RebootRequired を返す）
   ↓
wu_state.json（LoopCount, MaxLoops, InstalledKBs, FailedKBs, StartTime, RebootSec, AutoLogon）
   ↓ RebootRequired && InstalledCount > 0 && LoopCount < MaxLoops:
   Set-WindowsUpdateAutoLogon（autologon_list.csv から該当 user を引いて
                              AutoLogonCount = MaxLoops*2 で書く。CBS 消費分）
   Register-FabriqRunOnce（Profile __RESTART__ と同じ entry）
   Invoke-CountdownRestart
   ↓ 再起動 → RunOnce → Fabriq.exe 自動起動
   ↓ main.ps1 冒頭で Test-Path wu_state.json → Invoke-WindowsUpdateLoop 再入
   ↓ LoopCount++ で次の pass、繰り返し
   ↓ RebootRequired が消える or LoopCount >= MaxLoops:
   Clear-WindowsUpdateAutoLogon（registry 残留 password / count を消す）
   Show-WindowsUpdateSummary
   wu_completed.json（finalize 用 metadata）を出力
```

**Phantom KB 検出**: `InstalledKBs` で同じ KB が 3 回以上現れたら `skipKBs` に入れて以後インストール対象から除外（再出現する不良 KB 対策）。

`MaxRebootLoops` / `RebootCountdownSeconds` / `AutoLogonEnabled` は `windows_update_list.csv` から読む（デフォ 5 / 15 / true）。

---

## Complete-ProfileExecution（finalize パイプライン）

profile 実行末尾の集約処理。Linear と FlexProfile `[Complete]` の両方で呼ばれる。

```
Mode='Auto' or 'Manual'
   ↓
1. Export-ExecutionHistory       ── execution_history.csv → logs/history/history_export_*.csv
                                  + evidence/{base}/export_history/ にもコピー
2. Export-HtmlChecklist           ── evidence/{base}/checklist/checklist_*.html を生成
                                    プロファイル定義 vs 実行結果のクロスマップ
                                    ネットワーク照合 / プリンタ照合 / Windows License /
                                    BitLocker ステータスを含む完全な audit レポート
3. Mode='Auto'（Linear）:
   - log_uploader を silent mode で発火（バックグラウンド送信）
   - view_report.ps1 で HTML を最後に開く
   Mode='Manual'（Flex [Complete]）:
   - 同上だが UI 順序が異なる（operator が成果物確認後に発火するため）
```

HTML チェックリストの中身は §07_evidence_history.md に詳述。


<!-- ============================================================ -->
# === kernel/04_csv_encryption.md ===
<!-- ============================================================ -->

# CSV 駆動 + 暗号化アーキテクチャ

fabriq 全体は CSV 駆動で動く。ホスト情報・モジュール設定・プロファイル定義・カテゴリマスタ・作業者・ログ配送先 — すべて CSV。機密値はフィールド単位で暗号化して同じ CSV に共存させる。

---

## 暗号化仕様（全モジュール共通）

| 項目 | 値 |
|---|---|
| 鍵導出 | PBKDF2-HMAC-SHA256, **100,000 iterations**, 固定ソルト `fabriq-fixed-salt-2024` |
| 暗号アルゴリズム | AES-256-CBC, PKCS7 padding |
| エンコード | UTF-8（平文）→ Base64（暗号文） |
| プレフィックス | `ENC:<Base64>` |
| 鍵長 | 32 bytes (AES-256) |
| IV 長 | 16 bytes (AES block size) — PBKDF2 から派生（鍵と同じ KDF stream を再利用） |

### `Unprotect-FabriqValue` の実装ポイント（common.ps1 §445）

- `ENC:` プレフィックスが無い値はそのまま返す（`Import-ModuleCsv` で混在容認）
- `Rfc2898DeriveBytes` で key と iv を**一回の KDF 流れから連続取得**（C# `CryptoPoC` と完全互換）
- `MemoryStream` → `CryptoStream` → `StreamReader` で UTF-8 復号
- すべての CryptoStream / Aes / KDF オブジェクトは `Dispose` 確実

### パスフレーズ検証トークン（`kernel/txt/passphrase_verify.txt`）

- 中身: `ENC:<Base64>` 形式で、平文 `"surkitinisme"` を当該パスフレーズで暗号化したもの
- Fabriq Studio がパスフレーズ初回設定時に生成
- `Test-MasterPassphrase` が起動時にこのトークンを復号 → 結果が `"surkitinisme"` ならパスフレーズ正解
- このファイルが無いと Fabriq は起動できない（exit 1）。Studio が必ず先行する設計

### Resume 時のパスフレーズ持ち回り（DPAPI）

再起動跨ぎは AES とは別経路：

- `Protect-PassphraseForResume`: パスフレーズを **DPAPI LocalMachine** で暗号化して Base64 文字列化、`resume_state.json` の `ProtectedPassphrase` フィールドに格納
- `Unprotect-PassphraseFromResume`: 同マシンの DPAPI で復号
- LocalMachine スコープなのでマシンが変わると復号できない（盗難 PC で resume が走らない安全弁）
- DPAPI 復号失敗時は手動パスフレーズ再入力にフォールバック（最大 3 回）

---

## Import-ModuleCsv（モジュールが必ず通す統合パイプライン）

```powershell
$rows = Import-ModuleCsv -Path $listCsv `
                        -FilterEnabled `
                        -RequiredColumns @("Enabled","Path","Type") `
                        -Segment $env:FABRIQ_SEGMENT
```

### 4 段階パイプライン

```
1. Import-CsvSafe
   ├── Test-Path 不在 → Show-Error + return $null
   └── Import-Csv -Encoding Default で文字化け回避（PS5.1 のシステム既定 = SJIS / UTF-8 自動判別）

2. 透過復号
   ├── 各 row の各 prop を走査
   ├── 文字列値で `ENC:` 始まりなら Unprotect-FabriqValue
   └── 失敗時は Show-Warning + 元の暗号文のまま温存（モジュール側で判断可能）

3. RequiredColumns 検証（指定時のみ）
   ├── 1 行目の PSObject.Properties.Name で列存在チェック
   └── 欠落あれば Show-Error + return $null（モジュール先頭で空 return パターン）

4. FilterEnabled + Segment フィルタ
   ├── Enabled='1' で絞る（FilterEnabled 指定時のみ）
   ├── Segment 列があれば $env:FABRIQ_SEGMENT と厳密マッチ（空 vs 空もマッチ）
   └── ヒット 0 件なら Show-Skip + return @() （モジュールは 0 件を Skipped として返す）
```

### Segment の使い分け例

```csv
# wallpaper_list.csv
Enabled,Segment,WallpaperPath
1,office,\\share\wallpapers\office.png
1,home,\\share\wallpapers\home.png
1,,                                       ← Segment 空 = "デフォルトグループ"
```

```csv
# profile.csv
Order,ScriptPath,Enabled,Description,Segment
10,standard/wallpaper_config/wallpaper_config.ps1,1,オフィス用壁紙,office
20,standard/wallpaper_config/wallpaper_config.ps1,1,ホーム用壁紙,home
30,standard/wallpaper_config/wallpaper_config.ps1,1,既定壁紙,
```

3 行とも同じモジュールを呼ぶが、`FABRIQ_SEGMENT` 経由で異なる `_list.csv` 行に分岐する。

---

## CSV エンコーディング規約

| ファイル | エンコード | 改行 | 備考 |
|---|---|---|---|
| `kernel/csv/*.csv` | UTF-8 BOM | CRLF | hostlist / workers / categories / log_destinations / manifesto |
| `modules/*/module.csv` | UTF-8 BOM | CRLF | カーネルが `Import-Csv -Encoding Default` で読む |
| `modules/*/_list.csv` | UTF-8 BOM | CRLF | 各モジュール固有の設定 CSV |
| `modules/*/preset.csv` | UTF-8 BOM | CRLF | Studio 用ドロップダウン UI 定義 |
| `profiles/*.csv` | Default (SJIS or UTF-8 BOM) | CRLF | `Import-Csv -Encoding Default` で読み書き |

PS5.1 + Windows の制約で「日本語含む CSV は UTF-8 BOM が安全」が運用ルール（feedback memory `feedback_ps1_utf8_bom`）。

---

## CSV ファイルの分類（更新オーバーレイから見た 3 区分）

### 1. **Framework CSV**（テンプレ → ターゲットへ overlay 対象）

| ファイル | 役割 | 配置 |
|---|---|---|
| `module.csv` | モジュールメニュー定義（MenuName, Category, Order, Script, Enabled） | `modules/{type}/<name>/module.csv` |
| `preset.csv` | Studio 用ドロップダウン UI（Column, Value, Label） | `modules/{type}/<name>/preset.csv` |

`dev/framework_overlay_rules.json` の `moduleCsvWhitelist` に列挙。

### 2. **Site-Specific CSV**（絶対保護、上書き禁止）

| ファイル | 中身 | 保護理由 |
|---|---|---|
| `kernel/csv/hostlist.csv` | 対象 PC マスタ（`ENC:` 暗号化済の機密フィールド含む） | 顧客固有 |
| `kernel/csv/workers.csv` | 作業者マスタ | 顧客固有 |
| `kernel/csv/log_destinations.csv` | ログ配送先 + 認証情報（`ENC:`） | 顧客固有 |
| `modules/*/_list.csv` | 各モジュールの設定データ（reg_template, app_config 等） | キッティング案件ごとに作る |
| `profiles/*.csv` | 実行プロファイル | 案件ごとに作る |

`framework_overlay_rules.json` の `excludeFilesKernelLevel` + `excludeDirsRecursive` ("profiles") + ホワイトリスト除外で守られる。

### 3. **Runtime Artifact**（実行時生成、配備からは除外）

| ファイル | 役割 |
|---|---|
| `kernel/json/status.json` | Status Monitor のライブ状態（atomic write） |
| `kernel/json/session.json` | 現セッション情報 |
| `kernel/json/resume_state.json` | 再起動跨ぎの状態スナップショット |
| `kernel/json/art_pulse.txt` | Show-* が +1 する鼓動カウンタ |
| `kernel/json/skip_request.flag` | async モジュール強制スキップ要求 |
| `kernel/txt/passphrase_verify.txt` | パスフレーズ検証トークン |
| `kernel/txt/silence.flag` | ART 演出抑制 flag |

---

## hostlist.csv の構造（運用上もっとも重要）

```csv
AdminID,OldPCName,NewPCName,
EthernetIP,EthernetSubnet,EthernetGateway,
WifiIP,WifiSubnet,WifiGateway,
DNS1,DNS2,DNS3,DNS4,
Pin,
Printer1Name,Printer1Driver,Printer1Port,
Printer2Name,...,Printer10Port
```

- `AdminID`: 管理 ID（一級識別子。実行履歴 / HTML チェックリスト / Restore-ExecutionHistory のフィルタキー）
- `OldPCName` / `NewPCName`: 旧 PC 名 / 新 PC 名（hostname_config が NewPCName へリネーム）
- 機密フィールドは Studio で `ENC:<Base64>` に暗号化可能（Pin / DNS / 各 Printer 等）
- ホスト選択時 `Set-SelectedHostEnvironment` がすべて env vars へ流し込み（ENC は復号して入る）

---

## 暗号化設計の哲学

> 「**鍵を分散させない、ファイルベースの透過復号で鍵管理を 1 つに集約**」

- 鍵は単一のマスターパスフレーズのみ（PBKDF2 で派生）
- パスフレーズの検証はトークン 1 つ（`passphrase_verify.txt`）
- すべての CSV は同じ鍵で読まれるため、operator の心理的負荷が小さい
- Studio が暗号化 / 復号 / 検証トークン生成を一括で担当する → fabriq 本体はパスフレーズを「使う」だけ
- 固定ソルトの選択は意図的（鍵をマスターパスフレーズだけに依存させ、複数 PC 間で同じ ENC 値が同じ平文に復号できるようにするため。Salt をユニークにすると配布フェーズで矛盾が起きる）

セキュリティ詰めの甘さの棚卸しは `project_crypto_security_review.md` で別途追跡されている（A: 単独で直せる衛生 / B: Studio 連携要する format migration）。


<!-- ============================================================ -->
# === kernel/05_resume_restart.md ===
<!-- ============================================================ -->

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


<!-- ============================================================ -->
# === kernel/06_status_monitor.md ===
<!-- ============================================================ -->

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


<!-- ============================================================ -->
# === kernel/07_evidence_history.md ===
<!-- ============================================================ -->

# エビデンス・実行履歴・HTML チェックリスト

fabriq は「**やった証拠を残す**」ことを業務契約として担保する。すべてのモジュール実行はスクリーンショット PNG + 実行履歴 CSV 行 + HTML チェックリストへ反映される。

---

## エビデンスベースパス（unified evidence directory）

`Initialize-EvidenceBasePath` がセッション開始時に決定：

```
.\evidence\{Timestamp}_{PCName}_{Serial}_evidence\evidence\
```

例:
```
.\evidence\2026_05_06_103045_NEW-PC-01_T2NXCV06Y22208C_evidence\evidence\
```

中身（サブディレクトリ）:

```
{base}/evidence/
├── auto_capture/         ── Capture-ScreenEvidence の自動 PNG
├── gyotaku/              ── Save-Screenshot の手動 PNG（Status Monitor のボタン経由）
├── checklist/            ── Export-HtmlChecklist の HTML レポート
├── export_history/       ── Export-ExecutionHistory のフルダンプ CSV
└── pc_information/       ── evidence_config モジュールの収集結果（22 セクション）
    └── {date}_{uid}_{pc}/
        ├── 01_SystemInfo.txt
        ├── ...
        ├── 22_OfficeLicense.txt
        └── manifest.json    ── EVIDENCE_MANIFEST.md 公開契約に基づく
```

`SELECTED_NEW_PCNAME` と Hardware Unique ID（BIOS Serial Number 優先、fallback で MAC Address）から命名。Windows パス禁止文字は `-` に sanitize。

---

## Capture-ScreenEvidence（自動スクリーンショット）

モジュール実行ごとに `Invoke-KittingScript` / `Invoke-BatchExecution` が後置で発火：

```powershell
Capture-ScreenEvidence -ModuleName $module.MenuName -Status $result.Status
```

### 動作

1. `[DPIUtil]::SetProcessDPIAware()` で物理解像度を取得（スケーリング無視）
2. `[Screen]::PrimaryScreen.Bounds` のサイズで Bitmap 作成
3. `Graphics.CopyFromScreen` で Primary Screen を描画
4. `evidence/{base}/auto_capture/{ts}_{ModuleName}_{Status}_{PCName}.png` に保存

### 命名規則

```
2026_05_06_103055_HostnameConfiguration_Success_NEW-PC-01.png
```

タイムスタンプ + モジュール名（path-invalid 文字は `_`）+ Status + PC 名。

### Save-Screenshot との違い

| 関数 | トリガ | 保存先 | DPI |
|---|---|---|---|
| `Capture-ScreenEvidence` | モジュール実行直後（自動） | `auto_capture/` | `SetProcessDPIAware()` 呼ぶ |
| `Save-Screenshot` | Status Monitor ボタン（手動） | `gyotaku/` | DPI 変更しない（GUI 破壊回避） |

`SetProcessDPIAware()` は不可逆で、呼ぶと WinForms ウィンドウが scaling 環境で縮む副作用がある。Status Monitor は GUI プロセスなので Save-Screenshot では呼ばない。

---

## execution_history.csv（実行履歴）

`logs/history/execution_history.csv` が永続的な追記 CSV。

### スキーマ（13 列、kernel 3.1.3 で `Order` 列追加）

| 列 | 中身 |
|---|---|
| `Timestamp` | `yyyy-MM-dd HH:mm:ss` |
| `KanriNo` | `$env:SELECTED_KANRI_NO` |
| `PCName` | `$env:SELECTED_NEW_PCNAME` |
| `ModuleName` | `module.csv` の MenuName（`[seg:office]` 等のサフィックス込み） |
| `Category` | カテゴリ |
| `Status` | `Success` / `Error` / `Skipped` / `Cancelled` / `Partial` / `Pending`（Flex の Reset State） |
| `Message` | `New-ModuleResult` の `Message`（カンマ・改行は CSV escape） |
| `WindowsUser` | `WindowsIdentity.GetCurrent().Name` |
| `Worker` | session.json の `WorkerName` |
| `MediaSerial` | session.json の `MediaSerial` |
| `SessionID` | `$script:SessionID`（`yyyyMMdd_HHmmss`） |
| `Verified` | `True` / `False` / 空（Post-Apply Verification の結果） |
| `Order` | profile CSV の `Order`（マーカー / ad-hoc は空） |

### 関数群

| 関数 | 役割 |
|---|---|
| `Initialize-ExecutionHistory` | ディレクトリ生成 + 旧 path（`kernel/execution_history.csv`）からの一回切り migration + `.bak` 作成 |
| `Write-ExecutionHistory` | 1 行追記。書き込み失敗時は 100ms 間隔で 3 回リトライ |
| `Import-ExecutionHistory [-FilterKanriNo] [-Limit]` | 全行 read → KanriNo フィルタ → Timestamp 降順 → Limit 件 |
| `Restore-ExecutionHistory [-SessionIDFilter]` | `$script:ExecutionResults` を CSV から再構築（IsRestored=$true マーク） |
| `Show-ExecutionHistory` | コンソール表示用 |
| `Export-ExecutionHistory` | `logs/history/history_export_{ts}.csv` + `evidence/{base}/export_history/...csv` への dual export |

### Restore-ExecutionHistory の二モード

- **legacy mode**（無引数）: 同 KanriNo の最新 50 件 + `--- Current Session ---` separator を append
- **filter mode**（`-SessionIDFilter`）: 指定 SessionID のみ、件数無制限、separator なし。FlexProfile が batch 開始時に呼んで in-batch state refresh に使う

---

## Export-HtmlChecklist（HTML 監査レポート）

`Complete-ProfileExecution` が finalize 段で呼ぶ。`evidence/{base}/checklist/checklist_{ts}.html` を生成。

### 中身（セクション別）

#### 1. Meta Grid

- AdminID / OldPCName / NewPCName / Profile / Worker / Generated At / Elapsed
- HardwareUniqueId / MediaSerial

#### 2. Network Verification（期待値 vs 実際値）

`Get-CurrentPCInfo` で実 OS から取った値と SELECTED_* 期待値を **行ベースで比較**：

| 項目 | Expected | Actual | Match |
|---|---|---|---|
| PC Name | NEW-PC-01 | NEW-PC-01 | ✓ |
| Ethernet IP | 192.168.1.100 | 192.168.1.100 | ✓ |
| Eth Subnet | 255.255.255.0 | 255.255.255.0 | ✓ |
| Eth Gateway | 192.168.1.1 | 192.168.1.1 | ✓ |
| Wi-Fi IP | 192.168.5.20 | (none) | ✗ |
| DNS | 8.8.8.8, 8.8.4.4 | 8.8.8.8, 8.8.4.4 | ✓ |

DNS は **set-based 比較**（順序非依存、sort してから join）。

#### 3. Printer Cross-Check（3-way）

- **Expected**: hostlist.csv の `Printer<N>Name` 列
- **Actual**: `Get-Printer` の Network / IP port printers
- **Extra**: Actual にあるが Expected に無い（人間が手で入れた、または別経路で増えた）

#### 4. Windows License Status

WMI `SoftwareLicensingProduct` で `LicenseStatus` を取得し human-readable に変換：

```
0 = Unlicensed       (NG, red)
1 = Licensed         (OK, green)
2 = OOB Grace        (Partial, yellow)
3 = OOT Grace        (Partial, yellow)
4 = Non-Genuine Grace
5 = Notification
6 = Extended Grace
```

#### 5. BitLocker Status (C:)

`Get-BitLockerVolume -MountPoint C:` で `ProtectionStatus` / `VolumeStatus` を表示。

#### 6. Module Execution Checklist

profile CSV の `DefinedModules` 全件 vs `ExecutionResults` を Order でクロスマップ：

| Order | Description（CSV から） | MenuName | Status | Verified | Message |
|---|---|---|---|---|---|
| 10 | ホスト名設定 | Hostname Configuration | ✓ Success | ✓ PASS | Renamed to NEW-PC-01 |
| 20 | IP 設定 | IP Address Configuration | ✓ Success | ✓ PASS | DHCP→Static OK |
| 30 | ドメイン参加 | Domain Join | ⚠ Partial | — | 1 of 2 attempted |

`IsRestored=$true` の行も含めて選別（`Select-Object -Last 1`）し、再起動跨ぎでも最新結果が反映される。

### HTML 生成の特徴

- `System.Web.HttpUtility.HtmlEncode` で XSS 防御
- 純粋な inline CSS（外部依存なし、disconnected 環境でも閲覧可能）
- 印刷を意識した最小色彩 + 大きなフォント
- 色分け: Success=green / Partial=yellow / Skipped=gray / Cancelled=yellow / Error=red / Pending=blue / NotRun=gray

---

## evidence_config モジュール（pc_information 収集）

`modules/standard/evidence_config/` は fabriq 同梱の最大規模モジュール（v1.3.0+）。22 セクションのシステム情報を `pc_information/{date}_{uid}_{pc}/` 配下に出力し、最後に `manifest.json` を `EVIDENCE_MANIFEST.md` 契約に従って書く。

### 出力例

```
01_SystemInfo.txt              ── ComputerInfo, OSVersion, BIOS
02_HardwareInfo.txt            ── CPU, RAM, Disk, GPU
03_NetworkConfig.txt           ── ipconfig /all
...
14_ServerRolesFeatures.csv     ── Server OS のみ（Client は Skipped）
15_InstalledKBs.txt            ── Get-HotFix
...
21_Activation.txt              ── slmgr /xpr
22_OfficeLicense.txt           ── Office インストールあれば、なければ Skipped
manifest.json                  ── 全セクションの schemaVersion=1 manifest
```

詳細は §11_evidence_manifest_contract.md。

---

## log_uploader モジュール（外部共有への転送）

`modules/extended/log_uploader/` が Linear `[Execute Profile]` の finalize で auto-fire（`silent` mode）。FlexProfile では `[Complete]` ボタンで明示発火。

### 配送先設定（kernel/csv/log_destinations.csv）

| 列 | 意味 |
|---|---|
| Enabled | `1` で有効 |
| Name | 表示名 |
| Path | UNC パス or ローカル絶対パス |
| Username | UNC 認証用（任意、`ENC:` 暗号化可） |
| Password | UNC 認証用（任意、`ENC:` 暗号化可） |

### 配送内容

- `evidence/{base}/` 全体（auto_capture / gyotaku / checklist / export_history / pc_information）
- `logs/{ts}_{uid}_{hostname}.log`（Transcript）
- `execution_history.csv`

robocopy で送る。失敗してもキッティング自体は成功扱い（log 配送はベストエフォート）。


<!-- ============================================================ -->
# === kernel/08_async_execution.md ===
<!-- ============================================================ -->

# 非同期モジュール実行（__ASYNC__ + Runspace + Skip）

`__ASYNC__` マーカー以降のモジュールを別 PowerShell Runspace で実行し、Status Monitor からの **Skip ボタン** または **タイムアウト** で強制中断できる仕組み。kernel 2.1.0 で追加。

---

## なぜ async 実行が必要か

通常の `Invoke-SafeCommand` は同期実行で、モジュールが内部でハングすると fabriq 全体がブロックされる。長時間処理（domain join, winget install など）でこのリスクが顕在化したため、選択的に Runspace 実行に切り替える機構を導入した。

ただし「全モジュール async」は Runspace 起動コストとデバッグ性の問題があるため、**プロファイルで明示的に opt-in** する設計。

---

## kernel/json/async_config.json

```json
{
  "Enabled": true,
  "DefaultTimeoutSec": 0,
  "PollIntervalMs": 500,
  "SkipFlagPath": ".\\kernel\\json\\skip_request.flag"
}
```

| キー | 用途 |
|---|---|
| `Enabled` | `false` で `__ASYNC__` マーカーを silent ignore（kill switch） |
| `DefaultTimeoutSec` | `0`=無制限。`> 0` なら全 async モジュールにこの timeout を適用 |
| `PollIntervalMs` | Skip flag / completion / timeout チェックの polling 間隔（最小 50ms に clamp） |
| `SkipFlagPath` | Status Monitor が touch する flag ファイル絶対パス |

Kill switch を kernel 側に置いた理由: 緊急時に運用環境でも `Enabled=false` に書き換えるだけで全モジュールが従来の同期実行へ自動 fallback できる。

---

## Resolve-ProfileModules での解釈

```
profile CSV:
   Order, ScriptPath
   10,    standard/hostname_config/hostname_config.ps1
   20,    __ASYNC__                         ← この行は ValidModules に入らない
   30,    standard/winget_install/winget_install.ps1   ← _IsAsync=$true
   40,    standard/app_config/app_config.ps1            ← _IsAsync=$true（sticky）
   50,    __RESTART__
   60,    standard/reg_hklm_config/reg_hklm_config.ps1  ← _IsAsync=$true（profile 末尾まで継続）
```

`__ASYNC__` はプロファイル末尾までの sticky フラグ。途中で sync に戻すマーカーは無い（必要なら別プロファイルに分割）。

---

## Invoke-SafeCommandAsync の実装

`common.ps1` L1008、約 220 行。

### Runspace 立ち上げ

```powershell
$runspace = [runspacefactory]::CreateRunspace($Host)
$runspace.ApartmentState = "STA"           # WinForms 互換のため
$runspace.ThreadOptions  = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.Path.SetLocation($fabriqRoot)  # 相対パス整合
```

### Globals 注入

```powershell
$inject = @{
    FabriqMasterPassphrase = $global:FabriqMasterPassphrase
    AutoPilotMode          = $global:AutoPilotMode
    AutoPilotWaitSec       = $global:AutoPilotWaitSec
    FabriqTranscriptPath   = $global:FabriqTranscriptPath
    FabriqUniqueId         = $global:FabriqUniqueId
    FabriqSessionTimestamp = $global:FabriqSessionTimestamp
    FabriqEvidenceBasePath = $global:FabriqEvidenceBasePath
    FabriqEvidenceRootPath = $global:FabriqEvidenceRootPath
}
```

env vars は Process スコープなので Runspace に自動継承される。`$global:*` だけ明示注入。

### Wrapper Script Block

```powershell
$ps.AddScript({
    param($CommonPath, $ModuleScript, $Inject, $FabriqRoot)
    Set-Location -Path $FabriqRoot
    . $CommonPath                          # Show-Info etc を再ロード
    foreach ($key in $Inject.Keys) {
        Set-Variable -Name $key -Value $Inject[$key] -Scope Global -Force
    }
    $global:_LastModuleResult = $null
    $output = & $ModuleScript
    [PSCustomObject]@{
        _AsyncWrapper = $true
        Output        = $output
        LastResult    = $global:_LastModuleResult
    }
})
```

`_LastModuleResult` も pipeline と並列で回収するため、wrapper PSCustomObject 経由で親 Runspace に運ぶ。

### Monitor Loop

```powershell
$asyncHandle = $ps.BeginInvoke()
while (-not $asyncHandle.IsCompleted) {
    Start-Sleep -Milliseconds $pollMs

    if (Test-Path $skipFlagPath) {
        Remove-Item $skipFlagPath -Force
        $interrupted = $true
        $interruptReason = "Skip"
        try { $ps.Stop() } catch { }       # Runspace 強制停止
        break
    }

    if ($TimeoutSec -gt 0 -and ((Get-Date) - $startTime).TotalSeconds -ge $TimeoutSec) {
        $interrupted = $true
        $interruptReason = "Timeout"
        try { $ps.Stop() } catch { }
        break
    }
}
```

中断時の result.Message:

- Skip: `"Module skipped by operator (async runspace stopped; system state may be incomplete)"`
- Timeout: `"Module exceeded timeout of {N}s (async runspace stopped; system state may be incomplete)"`

「system state may be incomplete」を文字列に含めることで、Skip 後の状態が部分適用である可能性を operator に明示する。

### 完走時の捕捉ロジック

```powershell
$wrapper = $wrappedOutput | Where { $_._AsyncWrapper -eq $true } | Select -First 1
$moduleResult = $wrapper.Output | Where { $_._IsModuleResult -eq $true } | Select -First 1
if (-not $moduleResult -and $null -ne $wrapper.LastResult) {
    $moduleResult = $wrapper.LastResult     # フォールバック
}
```

`Invoke-SafeCommand` と同じセマンティクスを保つ。

### クリーンアップ

```powershell
finally {
    $result.Duration = (Get-Date) - $startTime
    if ($null -ne $ps)       { try { $ps.Dispose() } catch { } }
    if ($null -ne $runspace) {
        try { $runspace.Close() } catch { }
        try { $runspace.Dispose() } catch { }
    }
}
```

Skip / Timeout / 例外いずれの経路でも Runspace は確実に解放される。

---

## Skip Flag Path Normalization

`SkipFlagPath` を絶対パスに正規化してから polling する：

```powershell
if (-not [System.IO.Path]::IsPathRooted($skipFlagPath)) {
    $skipFlagPath = Join-Path (Get-Location).Path $skipFlagPath
}
$skipFlagPath = [System.IO.Path]::GetFullPath($skipFlagPath)
```

理由: Status Monitor は別プロセスで cwd が異なる可能性があり、絶対パスに揃えないと Skip 要求を取りこぼす。長時間 polling の途中で fabriq 側 cwd が変わるケースもあるため、loop 開始前に固定する。

---

## Skip フローの全体像

```
[ Status Monitor process ]
    operator が [Skip] ボタンを押す
       ↓
    New-Item -Path .\kernel\json\skip_request.flag -Force
       ↓
[ Main process / Invoke-SafeCommandAsync ]
    polling loop が Test-Path で検出
       ↓
    Remove-Item で flag を消す（次の async 実行のために）
       ↓
    $ps.Stop() で Runspace を強制終了
       ↓
    result.Status = "Error" + 上記の Message
       ↓
    Add-ExecutionResult / Write-ExecutionHistory で記録
       ↓
    AutoPilot 中なら ErrorMode 分岐（skip / retry / ask）
```

---

## 設計上の注意

### 1. 同期実行を default にした理由

- Runspace 起動オーバーヘッド（数十〜数百 ms / module）
- デバッグ性低下（Runspace 内例外のスタックトレースが分かりにくい）
- 互換性のリスク（一部モジュールが `$Host.UI.PromptForChoice` などホスト依存 API を使うと STA Runspace で崩れる）

→ 「ハングが懸念される一部のモジュールに対して opt-in」が現実解。

### 2. AutoPilot skip/timeout 機能の rejected 経緯

過去に「AutoPilot 全体に skip/timeout を適用する」案が検討されたが、Runspace refactor が広範すぎる vs ハング occurrence が稀という比較で **rejected**（2026-04-22、project memory `project_autopilot_skip_rejected`）。`__ASYNC__` の選択的適用が現行の落とし所。

### 3. Skip した状態の不完全性

Skip された Runspace は強制停止のため、モジュールが書きかけたレジストリ・ファイル・ネットワーク設定が中途半端に残る可能性がある。fabriq はこれを「Error として記録、operator に通知、続行は ErrorMode 次第」と扱う。後続モジュールが影響を受ける場合は profile 設計時に skip 不可能な順序を組む（domain_join の前に hostname_config を絶対に通すなど）。


<!-- ============================================================ -->
# === kernel/09_versioning.md ===
<!-- ============================================================ -->

# バージョン管理（カーネル + モジュール SemVer + コンパチマトリクス）

fabriq は **カーネル API とモジュールを独立に SemVer 管理** する設計。Claude（実装担当）の手順制御によって整合性を担保する（ランタイムチェックは行わない）。

---

## 管理対象ファイル

| ファイル | 真のソース性 | 更新タイミング |
|---|---|---|
| `kernel/KERNEL_VERSION` | カーネル API SemVer の **唯一の真のソース** | 公開 API 変更時、`KERNEL_API.md` の §1〜§5 範囲に影響 |
| `kernel/KERNEL_API.md` | 公開 API サーフェスの明文化 | 公開 API 追加・削除・シグネチャ変更時（KERNEL_VERSION 昇格と同コミット） |
| `kernel/EVIDENCE_MANIFEST.md` | manifest.json 公開契約 | manifest schema 変更時（schemaVersion 昇格と同期） |
| `modules/{std,ext}/<name>/VERSION` | モジュール個別 SemVer | モジュール touched 時に SemVer 規則で昇格 |
| `modules/{std,ext}/<name>/REQUIRES_KERNEL` | モジュールが要求する最小カーネル API 版 | 新しい公開 API（KERNEL_API.md §1〜§5）への依存が増えた時のみ |
| `dev/template/VERSION` | 新規モジュール用テンプレ | 初版 `0.1.0`（開発中・未リリースの目印） |
| `dev/template/REQUIRES_KERNEL` | 新規モジュール用テンプレ | 現行カーネル版 |

`README.md` L1 / `kernel/common.ps1` L2 / `kernel/main.ps1` L3 の版表記は `KERNEL_VERSION` の `X.Y` に同期する（リリース時のみ）。

**全体を表す「ディストリビューション版」は持たない**。kernel と各モジュールが個別に進化し、外部更新ツール（fabriq_studio）が SemVer 比較で bundle 単位の置き換えを判断する。

---

## カーネル SemVer 影響判定（KERNEL_API.md §1〜§5 をベース）

| 影響 | 昇格 | 例 |
|---|---|---|
| **MAJOR** (X+1.0.0) | 公開 API 破壊的変更 | KERNEL_API.md 記載の関数削除・シグネチャ変更 / Profile CSV 必須列削除・改名 / `ModuleResult` フィールド削除・契約変更 / `SELECTED_*` 環境変数改名 / 特殊マーカー削除（kernel 3.0.0 で `__SHUTDOWN__` 等 4 種を削除した実例） |
| **MINOR** (X.Y+1.0) | 公開 API への後方互換な追加 | 公開関数追加 / Profile CSV 任意列追加（kernel 3.2.0 の `Group` 列） / 特殊マーカー追加（kernel 2.1.0 の `__ASYNC__`） / 新環境変数追加 / グローバル変数追加（kernel 3.1.0 の `$global:AutoConfirmMode`） |
| **PATCH** (X.Y.Z+1) | 内部実装のみの変更（公開 API 不変） | `Invoke-SafeCommand` 内部最適化 / `Resolve-ProfileModules` リファクタ / 状態 JSON スキーマ変更 / バグ修正 |

判定に迷ったら**大きい側に倒す**（CLAUDE.md ルール B）。

### KERNEL_API.md §8「API Version History」

各公開 API の導入バージョンが記録される。モジュールが `REQUIRES_KERNEL` を打鍵する際、使う API すべての導入版の最大値を取る運用。

例: `Show-Info`（2.0.0）, `Import-ModuleCsv`（2.0.0）, `Group` 列依存（3.2.0）を使うモジュール → Min Kernel API = **3.2.0**

ただし `Group` 列はプロファイル側のスキーマでありモジュールスクリプト単体には影響しないため、`REQUIRES_KERNEL` には基本含めない（kernel が解釈する）。

---

## モジュール SemVer 影響判定

| 影響 | 昇格 | 例 |
|---|---|---|
| **MAJOR** (X+1.0.0) | モジュール外部仕様破壊 | `_list.csv` 必須列削除 / モジュールの入出力契約変更 / preset.csv の意味変更 |
| **MINOR** (X.Y+1.0) | 後方互換な機能追加 | 新しい設定項目対応 / 新セグメント追加 / Post-Apply Verification 追加 |
| **PATCH** (X.Y.Z+1) | バグ修正・内部改良 | エッジケース修正 / ログ文言改善 / 内部リファクタ |

---

## ベースライン Seed 運用（CLAUDE.md ルール H）

**2026-04-23 以降、全モジュールに `VERSION=1.0.0` と `REQUIRES_KERNEL=2.0.0` が baseline として一斉 seed 済み**（`dev/seed_module_versions.ps1` による）。

### 歴史的経緯

当初は「Claude が初めて touched した時点で `1.0.0` を打鍵する（lazy seed）」方針だったが、fabriq_studio の update 機能を実装する過程で「両側 VERSION 欠損 = SKIP」が現実的に問題（古い target と現行 template で実際には差分があるのに検出不可）となり、一斉 seed に切り替えた。

`evidence_config` / `odt_config` など既に独自に進んでいた VERSION は保持されている。

`pianist` モジュールは v1.6.0（推進中の活発なモジュール、2189 行と他より大幅に大きい）。

### Seed の冪等実行

```powershell
pwsh ./dev/seed_module_versions.ps1 -DryRun
```

新規モジュール追加時に seed 漏れが疑われる場合の確認コマンド。idempotent で既存値は保持される。

---

## VERSION ファイル形式

```
1.2.0
```

- 1 行 `X.Y.Z` のみ
- 末尾改行 1 個（trailing newline 必須）
- `REQUIRES_KERNEL` も同形式

---

## 実装前宣言（CLAUDE.md ルール E）

カーネル / モジュールを修正する前に必ず出す：

```
【変更スコープ宣言】
- 対象: kernel / module:<name> / profile / doc
- 公開 API サーフェスへの影響: あり / なし
  （あり の場合: どの関数/変数/マーカー/スキーマ が変化するか）
- KERNEL_API.md 参照済み: yes
- 予想バージョン影響:
    kernel  : X.Y.Z → X.Y.Z（MAJOR / MINOR / PATCH / 変更なし）
    modules : <touched modules with predicted bumps>
- 既存モジュールへの波及: ゼロ / <具体リスト>
```

モジュール touched 時は追加で API 依存スキャン（ルール I）：

```
【モジュール API 依存スキャン】
- 使用公開関数（KERNEL_API.md §1）: <列挙>
- 使用公開グローバル（§2）: <列挙>
- 使用公開環境変数（§3）: <列挙>
- Profile CSV / 特殊マーカー依存（§4）: <列挙、なければ「なし」>
- ModuleResult 契約使用（§5）: yes / no
- Min Kernel API 版: X.Y.Z（KERNEL_API.md §8 で逆引き）
```

---

## 実装サマリ報告（CLAUDE.md ルール F）

実装完了報告に必ず含める：

```
【バージョン影響サマリ】
- kernel/KERNEL_VERSION : X.Y.Z → X.Y.Z+N（種別 / 理由）
- KERNEL_API.md の更新 : あり / なし
- touched modules :
    <module_name> : X.Y.Z → X.Y.Z+N（種別 / 理由）
- untouched modules : N/74（一切触っていないモジュール数）
- 配備方針 : kernel/ フォルダ差し替えのみで OK / モジュール X の更新も必要 / 全件再配布必要
```

---

## CHANGELOG 運用（Keep a Changelog 1.1.0 準拠）

`kernel/` / `modules/` / `apps/` / `commands/` / `profiles/` / `dev/template/` 配下のコードまたは CSV スキーマを変更した場合、**同じコミット内で必ず**：

1. `CHANGELOG.md` の `[Unreleased]` セクションに追記
2. カテゴリ（`Added` / `Changed` / `Deprecated` / `Removed` / `Fixed` / `Security`）を選ぶ
3. 行頭にコンポーネント名（`kernel/common.ps1:`, `modules/standard/<name>:`, `profiles:` 等）のプレフィックス
4. モジュールを更新した場合は該当モジュールの `VERSION` を昇格
5. 公開 API（KERNEL_API.md 範囲）に影響がある場合は同コミット内で `KERNEL_API.md` を更新

ドキュメントのみの修正（コメント、Guide.txt、README）は CHANGELOG 追記不要。

---

## リリース手順（ユーザー明示指示時のみ）

1. `kernel/KERNEL_VERSION` を新しい `X.Y.Z` に更新
2. `CHANGELOG.md` の `[Unreleased]` を `[X.Y.Z] - YYYY-MM-DD` に昇格、直上に空の `[Unreleased]` を再設
3. 以下 3 箇所の版表記を `X.Y` に同期:
   - `README.md` L1: `# Fabriq ver{X.Y}`
   - `kernel/common.ps1` L2: `# Easy Kitting Batch - Common Function Library v{X.Y}.Z`
   - `kernel/main.ps1` L3: `# Fabriq ver{X.Y} - Manifeste du Surkitinisme -`
4. `pwsh ./dev/check_version.ps1` で整合性確認
5. ユーザーに annotated タグコマンドを提示（`git tag -a kernel-vX.Y.Z -m "..."`、Claude 側では実行しない）

---

## 中央コンパチマトリクス（Layer 3、未実装）

`VERSION` + `REQUIRES_KERNEL` + `KERNEL_API.md` を全モジュール走査することで、将来 `kernel/MODULE_COMPAT.md` を自動生成する想定：

```
| Module | Version | Min Kernel API | Last Touch | Notes |
```

現時点では未実装。Layer 2 データ（`REQUIRES_KERNEL` ファイル）が十分に貯まった段階で `dev/build_compat_matrix.ps1` を実装し、コミット時 or 定期実行で再生成する運用に移行予定。

---

## 整合性チェックスクリプト

```
dev/check_version.ps1
```

`KERNEL_VERSION` と各ファイル版表記の整合を検証。非 0 終了したら版表記を揃えてからコミット。リリースフロー手順 4 で必ず通す。

---

## 現行版（2026-05-06 時点）

- `KERNEL_VERSION`: **3.2.2**
- `dev/template/VERSION`: `0.1.0`（次の新規モジュール開始版）
- `dev/template/REQUIRES_KERNEL`: 現行カーネルに同期
- 標準モジュール 60 件 / 拡張モジュール 14 件すべて baseline `1.0.0` / `2.0.0`（一部例外: pianist `1.6.0`, evidence_config `1.3.0`+, etc.）


<!-- ============================================================ -->
# === kernel/10_function_index.md ===
<!-- ============================================================ -->

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


<!-- ============================================================ -->
# === kernel/11_directory_layout.md ===
<!-- ============================================================ -->

# ディレクトリ構成全体図

fabriq リポジトリのファイルツリーを上から眺めた全体像。配備時の意味づけ・更新オーバーレイの境界・ランタイム生成物の配置を網羅する。

---

## トップレベル構造

```
fabriq/
├── Fabriq.exe                    ── C# エントリ（UAC 自動昇格 → main.ps1 起動）
├── Fabriq_IOS.exe                ── fabriq_ios サブプロジェクト用ランチャ（独立 SemVer）
├── Deploy.bat                    ── USB → 対象 PC へのデプロイヘルパ
├── README.md                     ── L1 に "# Fabriq verX.Y" の版表記
├── CHANGELOG.md                  ── Keep a Changelog 1.1.0 形式（[Unreleased] + リリース版）
├── CLAUDE.md                     ── Claude 開発時の絶対遵守ルール（テンプレ厳守、命名、SemVer）
├── LICENSE                       ── MIT License
├── THIRD_PARTY_NOTICES.md        ── 7-Zip 25.01 (LGPL) 等のサードパーティライセンス
│
├── kernel/                       ── ★カーネル（更新オーバーレイの中核）
├── modules/                      ── ★モジュール群（独立 SemVer）
├── apps/                         ── ★GUI ツール群（kernel bundle と同期）
├── commands/                     ── ★ユーティリティコマンド（kernel bundle と同期）
├── profiles/                     ── ★site-specific プロファイル（overlay 絶対保護）
├── dev/                          ── ★開発ツールチェーン（kernel bundle と同期）
│
├── evidence/                     ── (runtime) エビデンス出力先
└── logs/                         ── (runtime) ログ出力先
```

---

## kernel/

```
kernel/
├── KERNEL_VERSION                ── 3.2.2（カーネル API SemVer の真のソース、1 行）
├── KERNEL_API.md                 ── 公開 API サーフェス（§1〜§11）
├── EVIDENCE_MANIFEST.md          ── manifest.json 公開契約（schemaVersion=1）
├── common.ps1                    ── 90+ 共通関数ライブラリ（4371 行）
├── main.ps1                      ── メインスクリプト（1913 行、FlexProfile sub-loop / WU loop 含む）
│
├── csv/                          ── マスタ CSV（site-specific は overlay 除外）
│   ├── categories.csv            ── カテゴリ + 表示順 (framework)
│   ├── hostlist.csv              ── 対象 PC マスタ (site-specific, ENC 暗号化対応)
│   ├── workers.csv               ── 作業者 (site-specific)
│   ├── log_destinations.csv      ── ログ配送先 (site-specific, ENC)
│   └── manifesto.csv             ── マニフェスト本文 (framework, 演出)
│
├── json/                         ── ランタイム状態（overlay 除外）
│   ├── status.json               ── (runtime) Status Monitor のライブ状態
│   ├── session.json              ── (runtime) 現セッション情報
│   ├── resume_state.json         ── (runtime) 再起動跨ぎ状態（v1/v2 schema）
│   ├── async_config.json         ── (framework) __ASYNC__ Runspace 制御パラメータ
│   ├── art_pulse.txt             ── (runtime) 動作鼓動カウンタ
│   └── skip_request.flag         ── (runtime) async モジュール強制スキップ要求
│
├── ps1/                          ── カーネルサブスクリプト
│   ├── status_monitor.ps1        ── 別プロセス WinForms モニタ
│   ├── view_report.ps1           ── HTML チェックリスト単体ビューア
│   ├── manifesto.ps1             ── マニフェスト表示 GUI
│   └── art_display.ps1           ── ART 演出（status_monitor に統合済）
│
└── txt/                          ── テキストアセット（site-specific は overlay 除外）
    ├── passphrase_verify.txt     ── (site-specific) パスフレーズ検証トークン（Studio 生成、必須）
    ├── art_sentences.txt         ── (framework) ART pulse 表示文
    └── silence.flag              ── (site-specific) 演出抑制 flag（存在チェックのみ）
```

---

## modules/

```
modules/
├── standard/                     ── 標準モジュール群（60 件）
│   └── <name>/
│       ├── module.csv            ── (framework) MenuName, Category, Order, Script, Enabled
│       ├── preset.csv            ── (framework, optional) Studio 用ドロップダウン UI 定義
│       ├── <name>.ps1            ── 実行スクリプト本体（dev/template ベース）
│       ├── <other>.ps1           ── 補助スクリプト（_install / _uninstall / _backup / _restore 等）
│       ├── <name>_list.csv       ── (site-specific) 設定データ
│       ├── Guide.txt             ── (framework) 使い方ガイド（日本語）
│       ├── VERSION               ── (framework) モジュール SemVer（1 行 X.Y.Z）
│       └── REQUIRES_KERNEL       ── (framework) 要求最小カーネル版（1 行 X.Y.Z）
│
└── extended/                     ── 拡張モジュール群（15 件）
    └── <name>/                   ── 同上
```

### モジュール件数

- Standard: 60 件
- Extended: 15 件（README には 14 とあるが、現状は pianist 含む 15 件）
- 合計: 75 モジュール

---

## apps/

`fabriq_operator` を中心に、kernel bundle と同期する GUI サブプロジェクト群。

```
apps/
├── fabriq_operator/              ── ★メインダッシュボード GUI（main.ps1 から . source）
│   ├── fabriq_operator.ps1
│   └── lib/
│       ├── theme.ps1             ── 色・サイズ定数
│       ├── session_form.ps1      ── 起動時の worker / host / passphrase 一括入力
│       ├── dashboard_form.ps1    ── メインダッシュボード（タブ + ボタン）
│       ├── flex_dashboard.ps1    ── FlexProfile ダッシュボード（Groups バー含む）
│       ├── quickactions_dialog.ps1
│       └── apps_dialog.ps1       ── FabriqApps 起動ダイアログ
│
├── fabriq_ios/                   ── ★Cisco IOS 風シェル（独立 SemVer の "art" project）
│   ├── fabriq_ios.ps1
│   ├── VERSION                   ── 独立 SemVer
│   ├── SPEC.md
│   ├── README.md
│   ├── data/
│   │   ├── version_banner.txt
│   │   ├── syslog_messages.csv
│   │   ├── help_text.csv
│   │   └── module_categories.json
│   ├── lib/
│   │   ├── parser.ps1 / dispatch.ps1 / completer.ps1 / prompt.ps1
│   │   ├── shell_state.ps1 / help.ps1 / syslog.ps1
│   │   ├── modes/                ── user_exec / privileged_exec / global_config / interface_config / module_config
│   │   └── commands/             ── enable_disable / show / hostname / interface / ip_address / categories / module
│   └── tests/                    ── _phase3_smoke.ps1 .. _phase9_smoke.ps1, parser/prompt/completer の unit tests
│
├── csv_editor/                   ── 汎用 CSV 編集 GUI
├── system_launcher/              ── Windows 設定ショートカット
├── bloatware_exporter/           ── インストール済アプリ一覧 export
├── desktop_icon_backup_app/      ── デスクトップアイコン backup
├── local_user_setup/             ── ローカルユーザー作成 GUI
├── storeapp_editor/              ── storeapp_list.csv 編集
└── winget_gui/                   ── winget 操作 GUI
```

---

## commands/

```
commands/
├── gpupdate_command.ps1          ── gpupdate /force ラッパー
├── temp_command.ps1              ── 一時的なコマンド枠
├── explore_restart_command.ps1   ── Explorer 再起動（ad-hoc）
├── diag_crypto.ps1               ── 暗号化機能の診断
├── get_evidence.ps1              ── 現セッションのエビデンス収集
└── system_launcher.ps1           ── System Launcher 経由から呼ばれる
```

---

## profiles/

```
profiles/
├── Master_Pre01.csv              ── マスタ pre-config phase 1
├── Master_Pre02.csv              ── pre-config phase 2
├── Master_Config01.csv           ── マスタ config phase 1
├── Master_Config02.csv           ── phase 2
├── Master_Config03.csv           ── phase 3
├── Master_Config04.csv           ── phase 4
├── Custom Plan.csv               ── 顧客カスタマイズプラン
├── sysprep.csv                   ── Sysprep 実行用
├── _test_harness.csv             ── テストハーネス（test_harness_config 用）
├── _test_harness_async.csv       ── async モード検証用
└── easy_template/                ── EasyProfile（簡易プロファイル実行）
    ├── easyprofile.bat
    ├── easyprofile.ps1
    └── easyprofile.csv
```

profiles/ は overlay の **excludeDirsRecursive** で完全保護される（顧客カスタムが上書きされない）。

---

## dev/

```
dev/
├── template/                     ── 新規モジュールスケルトン
│   ├── _template_script.ps1      ── 7-step 正典スケルトン（Step 1..7 構造）
│   ├── _template_list.csv        ── 設定 CSV のテンプレ（Enabled, Path, Type 等）
│   ├── module.csv                ── テンプレ用 module.csv
│   ├── Guide.txt                 ── 使い方ガイドのテンプレ
│   ├── VERSION                   ── 0.1.0（開発中・未リリースの目印）
│   └── REQUIRES_KERNEL           ── 現行カーネル版に同期
│
├── framework_overlay_rules.json  ── 更新オーバーレイ契約（schemaVersion=1）
├── build_framework_patch.ps1     ── フレームワーク更新パッチ生成
├── seed_module_versions.ps1      ── 全モジュール VERSION/REQUIRES_KERNEL の baseline seed (idempotent)
├── check_version.ps1             ── KERNEL_VERSION と版表記の整合チェック
├── verify_comments_only.ps1      ── スクリプトコメント英語化検証
│
├── launcher/                     ── Fabriq.exe / Fabriq_IOS.exe の C# ソース
│   ├── README.md
│   └── (.csproj, .cs sources)
│
└── ico/                          ── アイコン素材
```

---

## evidence/（runtime）

```
evidence/
└── {Timestamp}_{PCName}_{SerialNumber}_evidence/
    └── evidence/
        ├── auto_capture/             ── Capture-ScreenEvidence の PNG（モジュール毎）
        ├── gyotaku/                  ── Save-Screenshot の手動 PNG（Status Monitor ボタン）
        ├── checklist/                ── Export-HtmlChecklist の HTML 監査レポート
        ├── export_history/           ── Export-ExecutionHistory のフルダンプ CSV
        └── pc_information/           ── evidence_config モジュールの収集結果
            └── {date}_{uid}_{pc}/
                ├── 01_SystemInfo.txt .. 22_OfficeLicense.txt
                ├── manifest.json     ── EVIDENCE_MANIFEST.md 公開契約準拠
                └── manifest.json.bak ── 前回分（1 世代保持）
```

evidence ディレクトリは overlay の **excludeDirsTopLevel** で除外（顧客固有の成果物として保護）。

---

## logs/（runtime）

```
logs/
├── {Timestamp}_{uid}_{hostname}.log  ── PowerShell Transcript（セッション毎）
└── history/
    ├── execution_history.csv         ── 全セッション通算の実行履歴
    ├── execution_history.csv.bak     ── 起動時自動 backup
    └── history_export_{ts}.csv       ── プロファイル完了時の export
```

logs/ も overlay 除外。

---

## .git / .claude（開発時のみ）

```
.git/                              ── Git リポジトリ
.claude/                           ── Claude 関連の cache / settings
```

両者とも overlay の **excludeDirsTopLevel** で除外。

---

## 配備パッケージとしての境界

| 区分 | 含まれるか |
|---|---|
| **Framework**（kernel bundle で overlay 対象） | `kernel/` (csv/json/txt の site-specific は除く), `apps/`, `commands/`, `dev/`, `Fabriq.exe`, `Deploy.bat`, README, CHANGELOG, CLAUDE.md, LICENSE |
| **Module Bundle**（per-module overlay） | `modules/{std,ext}/<name>/` (module.csv + preset.csv + scripts + Guide.txt + VERSION + REQUIRES_KERNEL のみ、`_list.csv` 等は除く) |
| **Site-Specific**（絶対保護） | `profiles/` 全体, `kernel/csv/{hostlist,workers,log_destinations}.csv`, `kernel/txt/passphrase_verify.txt`, `modules/*/_list.csv`, `silence.flag` |
| **Runtime**（除外） | `kernel/json/*`, `evidence/`, `logs/`, `.git/`, `.claude/` |

詳細は `contracts/overlay_contract.md` を参照。


<!-- ============================================================ -->
# === contracts/profile_csv_schema.md ===
<!-- ============================================================ -->

# Profile CSV スキーマ契約

`KERNEL_API.md §4` で公式宣言。fabriq の最重要契約のひとつ。プロファイル CSV はキッティングシナリオを宣言する DSL であり、kernel と fabriq_studio の双方が依存する。

---

## ファイル配置

```
profiles/<ProfileName>.csv
```

例: `profiles/Master_Pre01.csv`, `profiles/Custom Plan.csv`, `profiles/sysprep.csv`

エンコーディングは Default（PS5.1 の `Import-Csv -Encoding Default` で読める SJIS / UTF-8 BOM のいずれか）。改行は CRLF。

---

## 列定義

| 列 | 必須 | 用途 | 由来版 |
|---|---|---|---|
| `Order` | 必須 | 整数・昇順実行。実行履歴の一級識別子（同 MenuName が複数行ある場合の per-row state matching に使う） | 2.0.0 baseline |
| `ScriptPath` | 必須 | `{standard,extended}/<module>/<script>.ps1` 形式 or 特殊マーカー。区切りは `/` `\` どちらも可 | 2.0.0 baseline |
| `Enabled` | 必須 | `1`=実行 / `0`=スキップ | 2.0.0 baseline |
| `Description` | 任意 | プロファイル UI 表示用コメント。`__AUTOPILOT__` 行では `WaitSec=N` 形式で wait 秒指定 | 2.0.0 baseline |
| `Segment` | 任意 | `Import-ModuleCsv` の Segment フィルタ値として渡される。`<name>_list.csv` の Segment 列と厳密マッチ | 2.0.0 baseline |
| `ErrorMode` | 任意 | AutoPilot 時のエラー処理（空=ダイアログ確認 / `skip` / `retry` 最大 5 回） | 2.0.0 baseline |
| `Group` | 任意 | FlexProfile dashboard の Groups バー集約名。Linear `Execute Profile` は無視 | **3.2.0** |

---

## 行の例

```csv
Order,ScriptPath,Enabled,Description,Segment,ErrorMode,Group
10,__AUTOPILOT__,1,WaitSec=3,,,
20,standard/hostname_config/hostname_config.ps1,1,ホスト名設定,,,Network
30,standard/ipaddress_config/ipaddress_config.ps1,1,IP アドレス設定,,retry,Network
40,__RESTART__,1,再起動,,,
50,standard/reg_hklm_config/reg_hklm_config.ps1,1,レジストリ設定,office,skip,Tweaks
60,standard/reg_hklm_config/reg_hklm_config.ps1,1,レジストリ設定（home），home,skip,Tweaks
70,__AUTO_to_admin01__,1,管理者で AutoLogon 設定,,,
80,__ASYNC__,1,以降を Runspace 化,,,
90,standard/winget_install/winget_install.ps1,1,Winget アプリ,,,Apps
100,standard/evidence_config/evidence_config.ps1,1,Evidence 収集,,,Evidence
```

---

## 特殊マーカー（5 種、kernel 3.0.0 で 4 種を破壊的削除）

### 現行マーカー

| マーカー | 動作 | 由来版 |
|---|---|---|
| `__AUTOPILOT__` | 以降を AutoPilot 化（Y/N 自動承認 + 指定 wait 秒のモジュール間スリープ）。`Description` に `WaitSec=N` で wait 秒指定 | 2.0.0 |
| `__ASYNC__` | 以降を Runspace 実行に切り替え。Status Monitor の Skip ボタン or `async_config.json` の `DefaultTimeoutSec` で強制中断可能 | 2.1.0 |
| `__RESTART__` | Windows 再起動 → RunOnce 経由で resume | 2.0.0 |
| `__REEXPLORER__` | Explorer 再起動（HKCU レジストリ変更の即時反映等） | 2.0.0 |
| `__AUTO_to_<User>__` | `autologon_config` を該当 User で呼び出し | 2.0.0 |

### 削除済みマーカー（kernel 3.0.0 / MAJOR）

`__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__` の 4 種を削除。

#### 削除理由

実運用での参照ゼロ（`__PAUSE__` / `__STOPLOG__` / `__STARTLOG__`）または唯一の使用箇所も廃止済み（`__SHUTDOWN__`）。fabriq_studio のマーカーパレットでも既に除外されており、UX 上は事実上 deprecated だった。

#### 既存プロファイル互換

削除後のマーカーを含む旧プロファイルは `Resolve-ProfileModules` の `$invalidPaths` 経由で「module not found」warning として降格、kernel はクラッシュせず他モジュールの実行を継続する（**graceful degradation**）。

---

## マーカー単位の挙動詳細

### `__AUTOPILOT__`

- `Enabled=1` のときのみ効果発動（`IncludeDisabled` モードでも disabled 行は無視される、安全弁）
- ValidModules には入らず、戻り値の `AutoPilot` フラグを `$true` にする
- `Description` に `WaitSec=3` のような形式で待機秒指定（regex `WaitSec=(\d+)`）
- 効果範囲: プロファイル末尾まで（off にするマーカーは無い）

### `__ASYNC__`

- `Enabled=1` かつ `async_config.json` の `Enabled=true` のときのみ効果発動（kill switch）
- ValidModules には入らない
- 以降のすべての通常モジュールに `_IsAsync=$true` を attach（sticky）
- 効果範囲: プロファイル末尾まで

### `__RESTART__`

- ValidModules に `_IsRestart=$true` の専用 PSCustomObject として入る
- `Invoke-BatchExecution` の loop 内で検出されると：
  1. `Save-ResumeState` で現状態を json 保存
  2. `Register-FabriqRunOnce` で次回起動を予約
  3. `Invoke-CountdownRestart -Seconds 5` で再起動（return）
- 再起動後の resume で「Order > restart 行の Order」のモジュールから続行

### `__REEXPLORER__`

- ValidModules に `_IsReexplorer=$true` で入る
- 検出時: `Stop-Process explorer -Force` → 最大 15 秒待機 → 自動復活しなければ `Start-Process explorer.exe`
- HKCU レジストリ変更（reg_hkcu_config）後の即時反映に使う

### `__AUTO_to_<User>__`

- 正規表現 `^__AUTO_to_(.+)__$` でマッチ
- 抽出された User 名が `_AutoLogonUser` として attach
- `autologon_config` モジュールが内部で `$env:FABRIQ_AUTOLOGON_USER` を読んで該当 user を `autologon_list.csv` から探す
- MenuName は `[AUTO:User] AutoLogon Configuration` に書き換えられる

---

## Group 列セマンティクス（kernel 3.2.0+）

### 動作モデル

- 同一 `Group` 値の行群を FlexProfile dashboard の **Groups バー** で `[Run: <Group>]` ボタンとして集約
- 1 クリックで該当 Group の全モジュールを Order 昇順で `Invoke-BatchExecution` する
- 実行は `AutoPilot=$true` + `FinalizeOnComplete:$false`（完了は operator が `[Complete]` で手動）
- Group 内の Order 順序は保たれる
- 空文字列 / 列自体の欠落 = 「グループ無所属」

### Linear への影響

Linear `[Execute Profile]` は本列を参照しない（旧来挙動維持）。Linear から見れば `_Group` 属性が module オブジェクトに増えるだけで無害。

### __RESTART__ との関係（literal interpretation）

Group 跨ぎ間の `__RESTART__`（Group 値が異なる）は当該 Group 実行時には **skip** される。**literal interpretation**: Group 列が batch を厳密に決定する契約。

operator が RESTART を含めたい場合は明示的に Group 値を打つ：

```csv
Order,ScriptPath,...,Group
10,standard/hostname_config/...,Network
20,standard/ipaddress_config/...,Network
30,__RESTART__,Network          ← Network グループで RESTART を含める
40,standard/domain_join/...,Network
```

`[Run: Network]` クリックで 10 → 20 → 30 (RESTART) → 再起動 → 40 まで一気通貫。

---

## __RESTART__ + AutoPilot の組み合わせ

最も強力な組み合わせ。再起動跨ぎを含む長尺シーケンスを完全 unattended 化（operator 立ち会い前提だが介入不要）：

```csv
10,__AUTOPILOT__,1,WaitSec=3
20,standard/hostname_config/hostname_config.ps1,1,
30,standard/ipaddress_config/ipaddress_config.ps1,1,
40,__RESTART__,1,
50,standard/domain_join/domain_join.ps1,1,
60,__RESTART__,1,
70,standard/bitlocker_config/bitlocker_config.ps1,1,
80,standard/evidence_config/evidence_config.ps1,1,
```

各 `__RESTART__` で resume_state.json + RunOnce + AutoLogon（autologon_config 経由）が組み合わさり、operator の介入なしで再起動 → 自動ログオン → 続行 → 再起動 → 自動ログオン → 続行 → 完了 まで進む。

---

## エラー処理マトリクス（ErrorMode 列）

| ErrorMode | AutoPilot=on | AutoPilot=off |
|---|---|---|
| 空文字 / `ask` | `Show-AutoPilotErrorDialog`（Retry/Skip 選択） | エラー記録のみ、続行 |
| `skip` | warning 出力 → エラー記録のまま続行 | 同左（off でも適用） |
| `retry` | autoRetryCount を最大 5 回まで増やして再実行、超えたらエラー記録 | 同左 |

ErrorMode は AutoPilot 中に最も意味を持つ（無人運用での自動復旧）。skip / retry は AutoPilot off 時もそのまま動く（operator が立ち会っていても自動 retry が便利な場合に対応）。

---

## Segment フィルタ仕様

- profile CSV 行の `Segment` 値が `$env:FABRIQ_SEGMENT` に渡される
- モジュール内の `Import-ModuleCsv` が `Segment` 列を持つ `_list.csv` に対して **厳密一致**（empty vs empty も match）でフィルタ
- 同モジュールを設定値別に呼び分けるパターンに使う（同じスクリプト、別データ）

例: 上の例の `Order=50` と `Order=60` は同じ `reg_hklm_config.ps1` を走らせるが、`reg_hklm_list.csv` の `Segment=office` 行と `Segment=home` 行を別々に適用する。

UI 上の MenuName 表示も `Registry HKLM Configuration [seg:office]` のようにサフィックス付きになる。

---

## デフォルトプロファイルの自動生成

`profiles/` ディレクトリが空のとき、`Create-DefaultProfiles` が以下 2 つを生成：

### Basic Setup.csv

```csv
Order,ScriptPath,Enabled,Description
10,standard\hostname_config\hostname_config.ps1,1,Change Hostname
20,standard\ipaddress_config\ipaddress_config.ps1,1,IP Address Settings
30,standard\domain_join\domain_join.ps1,1,Domain Join
```

### Full Setup.csv

検出された全モジュールを Order 10, 20, ... で並べた巨大プロファイル。

これらはあくまで初期値で、operator が Studio で site-specific プロファイルを作るのが本来の運用。


<!-- ============================================================ -->
# === contracts/module_result.md ===
<!-- ============================================================ -->

# ModuleResult 契約

`KERNEL_API.md §5` で公式宣言。fabriq モジュールがカーネルへ実行結果を返すための統一スキーマ。**すべてのモジュールはこの契約に従って結果を返却する**ことが義務付けられている。

---

## 戻り値オブジェクトのフィールド

| フィールド | 型 | 必須 | 意味 |
|---|---|---|---|
| `_IsModuleResult` | `bool` | 必須（常に `$true`） | カーネル側のフィルタが pipeline から ModuleResult を取り出す際の識別マーカー |
| `Status` | `string` | 必須 | `Success` / `Error` / `Cancelled` / `Skipped` / `Partial` のいずれか（ValidateSet） |
| `Message` | `string` | 必須（空文字 OK） | 結果の人間可読なメッセージ。実行履歴 CSV / HTML チェックリストに表示される |
| `Details` | `array` | 任意 | 任意の詳細情報。実行履歴には入らないが pipeline 経由でテストハーネスが消費可能 |
| `Verified` | `Nullable[bool]` | 任意 | Post-Apply Verification 結果。`$null`=未検証 / `$true`=PASS / `$false`=FAIL |
| `Timestamp` | `DateTime` | 必須（自動付与） | `Get-Date` で生成時刻を打鍵 |

---

## 標準呼び出しパターン

### 単純な完了

```powershell
return (New-ModuleResult -Status "Success" -Message "Done")
```

### 検証付き完了

```powershell
return (New-ModuleResult -Status "Success" -Message "Hostname renamed to NEW-PC-01 (pending reboot)" -Verified $true)
```

### バッチ集計

```powershell
return (New-BatchResult -Success 3 -Skip 1 -Fail 0 `
                        -Title "Registry Configuration" `
                        -MessageSuffix "(verified by readback)" `
                        -Verified $true)
```

`New-BatchResult` は内部で：
1. `Show-Separator` + `$Title` + 集計を画面表示
2. `$Status` を自動判定（Fail=0 + Success>0 → `Success` / Success>0 + Fail>0 → `Partial` / Skip 全件 → `Skipped` / Fail のみ → `Error`）
3. Message を `"Success: 3, Skip: 1, Fail: 0 (verified by readback)"` に組み立て
4. `New-ModuleResult` を呼んで返す

### キャンセル（Y/N で N 押下）

```powershell
$cancelled = Confirm-ModuleExecution -Message "本当に実行しますか？"
if ($cancelled) { return $cancelled }   # 即 return（既に Cancelled の ModuleResult が入っている）
```

### スキップ（対象なし）

```powershell
$rows = Import-ModuleCsv -Path $listCsv -FilterEnabled -RequiredColumns @("Enabled","Path")
if ($null -eq $rows -or $rows.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}
```

### エラー（例外発生）

```powershell
try {
    # ... 処理 ...
    return (New-ModuleResult -Status "Success" -Message "Applied $($results.Count) settings" -Verified $true)
}
catch {
    Show-Error "Configuration failed: $_"
    return (New-ModuleResult -Status "Error" -Message "Exception: $($_.Exception.Message)")
}
```

---

## Status セマンティクス

| Status | 意味 | HTML 表示色 | beep 通知 |
|---|---|---|---|
| `Success` | 全件正常完了 | 緑 | なし |
| `Partial` | 一部成功・一部失敗（New-BatchResult が Success>0 + Fail>0 で自動付与） | 黄 | 2-tone beep |
| `Skipped` | 全件スキップ（対象が無い、または明示スキップ） | 灰 | なし |
| `Cancelled` | ユーザーが Y/N で N 押下 | 黄 | なし |
| `Error` | エラー発生（例外 or 失敗） | 赤 | 3-tone beep |

`Invoke-ErrorNotification` は `Error` / `Partial` のときコンソールを foreground に持ってきて beep する（operator が見落とさないように）。

---

## Verified セマンティクス（Post-Apply Verification）

| Verified | 意味 | 推奨ケース |
|---|---|---|
| `$null` | 未検証 / 検証不可 | 検証ロジックを実装していないモジュール、または再起動後にしか確認できない（sysprep, hostname pending reboot 等） |
| `$true` | PASS（適用後の OS 状態が期待値と完全一致） | 設定読み返しが冪等性チェックと同じパターンで成立する場合 |
| `$false` | FAIL（適用は受理されたが乖離あり） | 適用 API が成功を返したのに読み返すと値が違う場合（外部要因や OS バグの兆候） |

### 実装パターン

```powershell
# Step 5.5 (Post-Apply Verification): 適用後に状態を読み返して照合
$verifiedCount = 0
$failedCount = 0
foreach ($entry in $rows) {
    if (Test-RegistryValueMatch -Path $entry.Path -Name $entry.Name -Expected $entry.Value) {
        $verifiedCount++
    } else {
        $failedCount++
    }
}
$allVerified = ($failedCount -eq 0)
return (New-BatchResult -Success $appliedCount -Skip $skipCount -Fail 0 `
                        -Title "Registry HKLM Results" -Verified $allVerified)
```

### 検証除外モジュール（feedback memory `project_verification_exclusions`）

以下は「**誤 PASS リスク回避**」の観点で意図的に検証を実装しない：

- `acl_config`: ACL ツリーを完全に読み返すと膨大、サブセット検証だと false PASS の可能性
- `spi_config`: Default Profile への hive load 経由のためログイン後にしか反映されず、検証時点では確定できない
- `copyfile_config`: ファイルが存在することと内容が正しいことは別、ハッシュ検証は重い

これらは `-Verified` 引数を省略（`$null`）。Guide.txt にも検証非実装の理由が明示される。

---

## Pipeline Capture と Fallback

カーネルは `Invoke-SafeCommand` 内で 2 段階で ModuleResult を捕捉：

```powershell
$global:_LastModuleResult = $null    # クリア
$output = & $ScriptBlock              # モジュール実行
foreach ($item in @($output)) {
    if ($item -is [PSCustomObject] -and $item._IsModuleResult -eq $true) {
        $moduleResult = $item
    }
}
if (-not $moduleResult -and $null -ne $global:_LastModuleResult) {
    $moduleResult = $global:_LastModuleResult   # フォールバック
}
```

**Pipeline capture 失敗例**: モジュールが内部で `Write-Output` を雑に呼び出して pipeline が混線、PowerShell が暗黙の implicit return で複数オブジェクトを return した場合に発生。`New-ModuleResult` 自体が内部で `$global:_LastModuleResult = $resultObj` を行うため、フォールバックで救済される（防御的二重化）。

---

## モジュールスクリプトの典型構造（7 step）

`dev/template/_template_script.ps1` がベースとなる正典スケルトン。すべてのモジュールはこの構造を踏襲する：

```powershell
# Step 1: Initialization
. "$PSScriptRoot\..\..\..\kernel\common.ps1"
$ModuleName = "Module Name"

# Step 2: Header (Show-CategorySeparator + Show-Info)
Show-CategorySeparator $ModuleName
Show-Info "..."

# Step 3: Confirmation
$cancelled = Confirm-ModuleExecution
if ($cancelled) { return $cancelled }

# Step 4: Load CSV
$csvPath = Join-Path $PSScriptRoot "<name>_list.csv"
$rows = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @(...)
if ($null -eq $rows) { return (New-ModuleResult -Status "Error" -Message "CSV load failed") }
if ($rows.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled rows") }

# Step 5: Apply (loop)
$successCount = 0; $skipCount = 0; $failCount = 0
foreach ($row in $rows) {
    # idempotency check (skip if already applied)
    # apply
    # increment counter
}

# Step 5.5: Post-Apply Verification (optional but recommended)
$verifiedAll = $true
foreach ($row in $rows) { if (-not (verify $row)) { $verifiedAll = $false } }

# Step 6: Display summary (handled by New-BatchResult)

# Step 7: Return result
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Verified $verifiedAll)
```

カーネルが Step 6 / Step 7 の表示を内部で完結させる設計のため、モジュール側は集計値だけ用意すればよい。


<!-- ============================================================ -->
# === contracts/special_markers.md ===
<!-- ============================================================ -->

# 特殊マーカー（Special Markers）契約

profile CSV の `ScriptPath` 列に書ける特殊識別子。`Resolve-ProfileModules` がこれらを解釈してプロファイル全体の挙動を制御する。

---

## 現行マーカー（5 種、kernel 3.2.x）

### 1. `__AUTOPILOT__`（kernel 2.0.0〜）

```csv
Order,ScriptPath,Enabled,Description,Segment,ErrorMode,Group
10,__AUTOPILOT__,1,WaitSec=3,,,
```

**動作**:
- `Enabled=1` のとき `$global:AutoPilotMode = $true` を立てる
- `Description` 列に `WaitSec=N` 形式があれば `$global:AutoPilotWaitSec` を `N` に設定（regex `WaitSec=(\d+)`）
- 効果は **以降のすべてのモジュール**に及ぶ（プロファイル末尾まで）
- ValidModules には**入らない**（メタデータ抽出のみ）

**AutoPilot の意味**:
- `Confirm-Execution` / `Confirm-ModuleExecution` が自動 Y を返す（Y/N プロンプトをスキップ）
- `Wait-KeyPress` が即 return（Press-Enter 待機をスキップ）
- 各モジュール開始前に `AutoPilotWaitSec` 秒の inter-module wait
- Error/Partial 時に `_ErrorMode` 列で分岐（`skip` / `retry` / `ask`）
- `Show-AutoPilotErrorDialog` が `ask` モードで Retry/Skip ダイアログ表示

**注意**: AutoPilot は「**確認スキップ + auto-resume**」であり「完全無人」ではない。operator が脇で見ていて状況に応じて Esc できる前提（feedback memory `feedback_autopilot_wording`）。

### 2. `__ASYNC__`（kernel 2.1.0〜）

```csv
20,__ASYNC__,1,以降を Runspace 化,,,
```

**動作**:
- `Enabled=1` かつ `async_config.json` の `Enabled=true` のときのみ効果発動（kill switch）
- 以降のすべての通常モジュールに `_IsAsync=$true` を attach（sticky）
- `Invoke-BatchExecution` が `_IsAsync` を見て `Invoke-SafeCommandAsync`（Runspace + monitoring）を選ぶ
- ValidModules には入らない
- 効果範囲: プロファイル末尾まで（途中で sync に戻すマーカーは無い）

**意味**: モジュール内ハングや長時間処理に対して、Status Monitor の Skip ボタン or `DefaultTimeoutSec` で強制中断できるようにする。詳細は §08_async_execution.md。

### 3. `__RESTART__`（kernel 2.0.0〜）

```csv
40,__RESTART__,1,再起動,,,
```

**動作**:
- ValidModules に `_IsRestart=$true` の専用 PSCustomObject として入る（MenuName=`[RESTART]`, Category=`System`）
- `Invoke-BatchExecution` の loop が検出すると：
  1. `Save-ResumeState` で `kernel/json/resume_state.json` に状態保存
  2. `Register-FabriqRunOnce` で次回起動を予約
  3. `Add-ExecutionResult [RESTART] Success` で履歴記録
  4. `Invoke-CountdownRestart -Seconds 5` で再起動（loop から return）
- 再起動後の resume で「Order > restart 行の Order」のモジュールから続行

**前提**: `autologon_config` モジュールが事前に AutoLogon を組んでいないと、再起動後に手動ログオンが必要になる（無人運用が成立しない）。

### 4. `__REEXPLORER__`（kernel 2.0.0〜）

```csv
50,__REEXPLORER__,1,Explorer 再起動,,,
```

**動作**:
- ValidModules に `_IsReexplorer=$true` で入る
- `Invoke-BatchExecution` 検出時:
  ```powershell
  Stop-Process -Name explorer -Force
  最大 15 秒、1 秒間隔で Get-Process explorer の復活を待つ
  自動復活しなければ Start-Process explorer.exe を強制発火
  Add-ExecutionResult [REEXPLORER] Success/Error
  ```
- 結果は `$completedResults` に正しい Status で記録される（kernel 3.x で改善、それ以前は hardcoded "Success" だった）

**用途**:
- HKCU レジストリ変更（reg_hkcu_config）を即時反映する
- タスクバー設定 / Start Layout / Explorer 動作系の変更後に必須
- Active Setup / Startup Batch 系のリフレッシュ

### 5. `__AUTO_to_<User>__`（kernel 2.0.0〜）

```csv
70,__AUTO_to_admin01__,1,管理者で AutoLogon 設定,,,
80,__AUTO_to_user01__,1,ユーザーで AutoLogon 設定,,,
```

**動作**:
- 正規表現 `^__AUTO_to_(.+)__$` でマッチ、`$Matches[1]` を `$autoLogonUser` として抽出
- `autologon_config` モジュールを参照し、PSCustomObject を copy
- `_AutoLogonUser` 属性を attach
- MenuName を `[AUTO:User] AutoLogon Configuration` に書き換え
- `Invoke-BatchExecution` 実行直前に `$env:FABRIQ_AUTOLOGON_USER = $module._AutoLogonUser`、終了後に `$null` でクリア
- `autologon_config` 内で `$env:FABRIQ_AUTOLOGON_USER` を読んで `autologon_list.csv` から該当行を選択

**用途**:
- 同じプロファイル内で複数の AutoLogon 切替が必要なシナリオ（admin で WU → user で BitLocker など）
- profile に `autologon_config` を直書きすると user 引数が固定されるが、`__AUTO_to_<User>__` で動的選択可能

---

## 削除済みマーカー（kernel 3.0.0 で MAJOR 破壊削除）

### `__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__`

| マーカー | 旧動作 | 削除理由 |
|---|---|---|
| `__SHUTDOWN__` | プロファイル末尾でシャットダウン | 唯一の使用箇所（profile 末尾）が廃止済み、`Invoke-CountdownShutdown` 内部関数も削除 |
| `__PAUSE__` | profile 中で `Wait-KeyPress` を強制発火 | 実運用での参照ゼロ |
| `__STOPLOG__` / `__STARTLOG__` | Transcript の一時停止/再開 | 実運用での参照ゼロ |

### graceful degradation

旧プロファイルがこれらを含んでいても fabriq はクラッシュせず、`Resolve-ProfileModules` の `$invalidPaths` 経由で「module not found」warning として降格、他モジュールの実行は継続する。fabriq_studio のマーカーパレットでも既に除外されているため、新規 profile 作成時には誤って書けない。

---

## 同 Group 内 / Group 跨ぎでのマーカー挙動（kernel 3.2.0+）

FlexProfile Groups バーの `[Run: <Group>]` ボタン経由で実行する場合：

- **同 Group 内**: マーカー（`__RESTART__` / `__REEXPLORER__` / `__AUTOPILOT__` / `__ASYNC__` / `__AUTO_to_<User>__`）はすべて **そのまま実行**
- **Group 跨ぎ間のマーカー**: 当該 Group 実行時には **skip**（**literal interpretation**）

operator が RESTART を含めたい場合は明示的に Group 値を打つ：

```csv
10,standard/hostname_config/...,1,,,,Network
20,standard/ipaddress_config/...,1,,,,Network
30,__RESTART__,1,,,,Network          ← Network グループで RESTART を含める
40,standard/domain_join/...,1,,,,Network
```

`[Run: Network]` クリックで 10 → 20 → 30 (RESTART) → 再起動 → 40 まで一気通貫。

Linear `[Execute Profile]` は Group 列を無視して全マーカーを順序通り実行する（旧来挙動維持）。

---

## マーカーと `IncludeDisabled` スイッチ

`Resolve-ProfileModules -IncludeDisabled` モード（FlexProfile 専用）でも：

- `__AUTOPILOT__` / `__ASYNC__` の効果発動には `Enabled=1` が必須（disabled 行は無視）
- `__RESTART__` / `__REEXPLORER__` / `__AUTO_to_<User>__` は ValidModules に入るが `_IsCheckedDefault=$false` がつく（dashboard で初期 unchecked）

「**disabled marker 行が global state を勝手に flip しない**」安全弁。FlexProfile でプロファイル全体を可視化しても、Enabled=0 のマーカー行が誤動作することはない。

---

## マーカー選択の運用知見

| 場面 | 推奨マーカー組み合わせ |
|---|---|
| 全自動キッティング（再起動含む） | `__AUTOPILOT__` + 各モジュール + `__RESTART__` 数箇所 + `autologon_config` 事前配置 |
| 短時間モジュールが連続するセクション | 単純な順序付けのみ（マーカー不要） |
| 長時間で hang リスクある winget / WU 系 | `__ASYNC__` を冒頭に置いて以降を runspace 化 |
| Operator 立ち会い・確認しながら実行 | マーカーなし（CLI モード廃止後は Linear `[Execute Profile]` で個別 Y/N） |
| Site-specific 段階的実行 | FlexProfile + Group 列で論理グループ化 |
| HKCU レジストリ変更後の即時反映 | reg_hkcu_config の直後に `__REEXPLORER__` |

---

## マーカー追加・削除のバージョン影響

- **新マーカー追加**: 公開 API への後方互換な追加 → kernel **MINOR** 昇格（例: 2.0.0 → 2.1.0 で `__ASYNC__` 追加）
- **既存マーカー削除**: 破壊的変更 → kernel **MAJOR** 昇格（例: 2.x → 3.0.0 で 4 マーカー削除）
- **マーカーの動作変更**: 後方互換でない場合は MAJOR / 互換なら MINOR

`KERNEL_API.md §4` がマーカー仕様の真のソース。Studio のマーカーパレットは `KERNEL_API.md` を参照して動的構築すべき設計。


<!-- ============================================================ -->
# === contracts/host_environment.md ===
<!-- ============================================================ -->

# SELECTED_* / FABRIQ_* 環境変数契約

`KERNEL_API.md §3` で公式宣言。fabriq の最重要 IPC（プロセス内コミュニケーション）。`hostlist.csv` の選択行から流れた値が、すべてのモジュールに共通して見える形で配信される。

---

## §3.1 ホスト情報変数（hostlist.csv 由来）

`Set-SelectedHostEnvironment` がホスト選択時点で `$env:` に流し込み。`ENC:` 暗号化フィールドは **この時点で透過復号** される。

### 識別系

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_KANRI_NO` | `AdminID` | 管理 ID（実行履歴の一級識別子、Restore-ExecutionHistory のフィルタキー、HTML チェックリスト Meta） |
| `SELECTED_OLD_PCNAME` | `OldPCName` | 旧 PC 名（リネーム前の参照用） |
| `SELECTED_NEW_PCNAME` | `NewPCName` | 新 PC 名（hostname_config が適用先、エビデンスベースパス命名 / HTML / status.json で常時参照） |

### ネットワーク系

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_ETH_IP` | `EthernetIP` | イーサネット IPv4 |
| `SELECTED_ETH_SUBNET` | `EthernetSubnet` | イーサネットサブネットマスク |
| `SELECTED_ETH_GATEWAY` | `EthernetGateway` | イーサネットゲートウェイ |
| `SELECTED_WIFI_IP` | `WifiIP` | Wi-Fi IPv4 |
| `SELECTED_WIFI_SUBNET` | `WifiSubnet` | Wi-Fi サブネットマスク |
| `SELECTED_WIFI_GATEWAY` | `WifiGateway` | Wi-Fi ゲートウェイ |
| `SELECTED_DNS1` 〜 `SELECTED_DNS4` | `DNS1`..`DNS4` | DNS サーバ最大 4 件 |

ipaddress_config / network_profile_config / DNS 系モジュールが消費。

### セキュリティ系

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_PIN` | `Pin` | セットアップ時の PIN（cert_config / sysprep 等で参照、`ENC:` 暗号化推奨） |

### プリンタ系（10 スロット）

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_PRINTER_<N>_NAME` | `Printer<N>Name` | プリンタ名（N=1..10） |
| `SELECTED_PRINTER_<N>_DRIVER` | `Printer<N>Driver` | ドライバ名 |
| `SELECTED_PRINTER_<N>_PORT` | `Printer<N>Port` | ポート（`IP_192.168.1.50` / `192.168.1.50` / TCPIP_*）|

`printer_driver_config` の register subroutine が消費。HTML チェックリストの Printer Cross-Check では Expected vs Actual で照合。

---

## §3.2 プロファイル実行パラメータ

| 環境変数 | 由来 | 寿命 |
|---|---|---|
| `FABRIQ_SEGMENT` | profile CSV の `Segment` 列 | モジュール 1 件の実行スコープ（Invoke-BatchExecution が前後で setup/teardown） |
| `FABRIQ_AUTOLOGON_USER` | `__AUTO_to_<User>__` マーカー | 同上（モジュール実行直前に立て、終了後に `$null`） |
| `FABRIQ_WORKER_NAME` | session.json `WorkerName` | セッション全体（Initialize-Session で設定） |
| `FABRIQ_EVIDENCE_BASE` | `Initialize-EvidenceBasePath` | セッション全体（`$global:FabriqEvidenceBasePath` と同値、モジュール内 `Join-Path` 用） |

### `FABRIQ_SEGMENT` の動作詳細

profile CSV の `Segment=office` 行から呼ばれた場合：

```
Invoke-BatchExecution の loop:
   1. $env:FABRIQ_SEGMENT = "office"
   2. & $module.Script   ← この間中 $env:FABRIQ_SEGMENT が見える
   3. $env:FABRIQ_SEGMENT = $null
```

モジュール側 `Import-ModuleCsv` のデフォルト引数 `-Segment $env:FABRIQ_SEGMENT` で自動的にフィルタが効く。

### `FABRIQ_AUTOLOGON_USER` の動作詳細

`__AUTO_to_admin01__` マーカーから呼ばれた場合：

```
Invoke-BatchExecution の loop:
   1. $env:FABRIQ_AUTOLOGON_USER = "admin01"
   2. & $autologon_config.ps1   ← 内部で $env:FABRIQ_AUTOLOGON_USER を読み、autologon_list.csv から User=admin01 行を選択
   3. $env:FABRIQ_AUTOLOGON_USER = $null
```

---

## §2 公開グローバル変数（補足）

環境変数とは別に、`$global:` スコープの公開変数も存在：

| グローバル | 型 | 環境変数版 | 違い |
|---|---|---|---|
| `$global:FabriqMasterPassphrase` | string | (なし) | 環境変数化しないことで child process 漏洩を防ぐ。Runspace 注入時のみ明示転送 |
| `$global:AutoPilotMode` | bool | (なし) | プロセス内 flag（環境変数化すると spawn された script や子プロセスにも影響して混乱） |
| `$global:AutoPilotWaitSec` | int | (なし) | 同上 |
| `$global:AutoConfirmMode` | bool | (なし) | FlexProfile 単発実行用 (kernel 3.1.0+) |
| `$global:FabriqEvidenceBasePath` | string | `FABRIQ_EVIDENCE_BASE` | 環境変数版も提供（モジュール内 `Join-Path` のため） |

---

## 環境変数のライフサイクル

### セッション全体

```
ホスト選択（Show-SessionSetupForm or resume の Restore-HostEnvironment）
   ↓
Set-SelectedHostEnvironment → 全 SELECTED_* 設定
   ↓
Initialize-Session → FABRIQ_WORKER_NAME 設定
   ↓
Initialize-EvidenceBasePath → FABRIQ_EVIDENCE_BASE 設定
   ↓
... モジュール実行中、SELECTED_* / FABRIQ_WORKER_NAME / FABRIQ_EVIDENCE_BASE は不変 ...
   ↓
Reset-FabriqState（New Session / Refabriq）
   ↓
すべての SELECTED_* / FABRIQ_AUTOLOGON_USER / FABRIQ_EVIDENCE_BASE を null 化
```

### モジュール 1 件のスコープ

```
Invoke-BatchExecution の foreach:
   1. _AutoLogonUser があれば $env:FABRIQ_AUTOLOGON_USER を立てる
   2. _Segment があれば $env:FABRIQ_SEGMENT を立てる
   3. & $module.Script
   4. $env:FABRIQ_AUTOLOGON_USER = $null（クリア）
   5. $env:FABRIQ_SEGMENT = $null
```

### 再起動跨ぎ

```
__RESTART__ 検出
   ↓
Save-ResumeState
   ├── 全 SELECTED_* を hash table 化（HostEnvironment fields）
   ├── EvidenceBasePath を json 保存
   └── ProtectedPassphrase（DPAPI 暗号化）も保存
   ↓ 再起動 ↓
   ↓
main.ps1 が Restore-HostEnvironment で全 SELECTED_* を json から戻す
   ↓
EvidenceBasePath / SessionID / ProtectedPassphrase（DPAPI 復号 → $global:FabriqMasterPassphrase）も復元
```

---

## 環境変数の Runspace 継承

`__ASYNC__` で Runspace 実行に切り替わった場合：

- `$env:*` は Process スコープなので **Runspace に自動継承される**（child runspace が parent process と同じ environment block を共有）
- `$global:*` は Runspace のスコープが異なるため **明示注入が必要**

`Invoke-SafeCommandAsync` の `$inject` ハッシュテーブルが `$global:*` を Runspace の global スコープへ Set-Variable で注入する。

---

## 環境変数の運用ルール

### 1. モジュールは SELECTED_* を読み取り専用扱いとする

```powershell
$pcName = $env:SELECTED_NEW_PCNAME   # OK: 読み取り
$env:SELECTED_NEW_PCNAME = "..."     # NG: モジュールから書き換えてはならない
```

理由: 同セッション内の他モジュールへ漏れて副作用となる。書き換えはカーネルの `Set-SelectedHostEnvironment` / `Restore-HostEnvironment` のみ許容。

### 2. ENC: 復号は kernel 経由で透過完了している

モジュールは `Unprotect-FabriqValue` を直接呼ぶ必要がない（hostlist.csv 由来は env 配信時、`_list.csv` 由来は `Import-ModuleCsv` で復号される）。

### 3. 機密値の env 配信を最小限に保つ

`SELECTED_PIN` / `SELECTED_PRINTER_<N>_*` 等の機密項目は env に出すが、**子プロセスを spawn する際は意識的に環境変数を絞る**（明示的な `Start-Process` 呼び出し時の Environment 引数で）。

### 4. fabriq_ios サブシェルからの参照

`apps/fabriq_ios` の `show ip interface` 等のコマンドは `$env:SELECTED_ETH_IP` 等を参照する。fabriq_ios も SELECTED_* 環境変数を「IOS のシステム情報」として表示する設計のため、契約破壊は fabriq_ios の挙動も壊す。


<!-- ============================================================ -->
# === contracts/overlay_contract.md ===
<!-- ============================================================ -->

# 更新オーバーレイ契約（External Tool ↔ fabriq）

`KERNEL_API.md §9` で公式宣言された、**外部更新ツール**（代表: `fabriq_studio`、および `dev/build_framework_patch.ps1`）が消費する公開契約。fabriq 本体の再配布・in-place 更新を「site-specific データを保持したまま framework 側だけ差し替える」運用で成立させる。

---

## 真のソース: `dev/framework_overlay_rules.json`

`schemaVersion=1` の JSON 1 ファイルが契約の真のソース。外部ツールはこのファイルを読んでルールを解釈する。

```json
{
  "schemaVersion": 1,
  "description": "Framework overlay rules for fabriq...",

  "excludeDirsTopLevel": [".git", ".claude", "evidence", "logs"],
  "excludeDirsRecursive": ["profiles"],
  "excludeFilesKernelLevel": [
    ".gitignore",
    "kernel/csv/hostlist.csv",
    "kernel/csv/workers.csv",
    "kernel/csv/log_destinations.csv",
    "kernel/json/art_pulse.txt",
    "kernel/json/resume_state.json",
    "kernel/json/session.json",
    "kernel/json/skip_request.flag",
    "kernel/json/status.json",
    "kernel/txt/passphrase_verify.txt",
    "kernel/txt/silence.flag"
  ],
  "moduleCsvWhitelist": ["module.csv", "preset.csv"],

  "bundles": {
    "kernel": {
      "versionFile": "kernel/KERNEL_VERSION",
      "includePaths": [
        "kernel/", "apps/", "commands/", "dev/",
        "Fabriq.exe", "Deploy.bat",
        "README.md", "CHANGELOG.md", "CLAUDE.md", "LICENSE"
      ]
    },
    "module": {
      "pathPattern": "modules/{type}/{name}/",
      "versionFilePattern": "modules/{type}/{name}/VERSION",
      "requiresKernelFilePattern": "modules/{type}/{name}/REQUIRES_KERNEL",
      "typeValues": ["standard", "extended"]
    }
  }
}
```

---

## Bundle 定義

| Bundle | Version ファイル | 対象パス |
|---|---|---|
| **kernel** | `kernel/KERNEL_VERSION` | `kernel/`, `apps/`, `commands/`, `dev/`, `Fabriq.exe`, `Deploy.bat`, `README.md`, `CHANGELOG.md`, `CLAUDE.md`, `LICENSE` |
| **module:\<name\>** | `modules/{std,ext}/<name>/VERSION` | `modules/{std,ext}/<name>/`（ただし `moduleCsvWhitelist` 以外の CSV は除く） |

`apps/` / `commands/` / `dev/` は個別 `VERSION` を持たず、kernel bundle と同期して動く。

---

## Site-Specific の絶対保護

更新時に **絶対に上書きしない** 対象：

### 1. profiles/ 配下全ファイル

`Master_*.csv`, `Custom Plan.csv`, `sysprep.csv`, `_test_harness*.csv`, `easy_template/` 等すべて。プロファイル書式のアップデートが入っても既存を優先。

### 2. excludeFilesKernelLevel に列挙された kernel 配下ファイル

- `kernel/csv/hostlist.csv` / `workers.csv` / `log_destinations.csv`（顧客固有マスタ）
- `kernel/json/*.json` 全般（runtime artifact）
- `kernel/txt/passphrase_verify.txt`（site 固有のマスターパスフレーズ検証トークン）
- `kernel/txt/silence.flag`（運用 opt-out flag）

### 3. modules/**/*.csv のうち moduleCsvWhitelist 以外のもの

`_list.csv` ファミリ、`office_key.csv`, `license_key.csv`, `domain.csv` 等すべての site-specific 設定 CSV。

### 4. ランタイム成果物

`kernel/json/*.json`, `art_pulse.txt`, `skip_request.flag`, `passphrase_verify.txt`, `silence.flag` 等。

---

## SemVer 比較セマンティクス

bundle 単位でバージョンを比較し、以下のテーブルに従って判断：

| template VERSION | target VERSION | 期待動作 |
|---|---|---|
| `1.2.0` | `1.1.0` | **UPDATE**（overlay 実行） |
| `1.2.0` | `1.2.0` | SKIP（同版） |
| `1.2.0` | `1.3.0` | SKIP（target 側が新しい。ツールによっては警告表示推奨） |
| `1.2.0` | VERSION ファイル欠損 | UPDATE（target を lazy seed） |
| なし | `1.0.0` | SKIP（template 側に VERSION 未打鍵） |
| なし | なし | SKIP |

バージョン文字列は `^(\d+)\.(\d+)\.(\d+)$` 形式。pre-release / build metadata は現行不使用。

---

## REQUIRES_KERNEL 事前チェック

モジュール bundle を overlay する前に、template 側のそのモジュールの `REQUIRES_KERNEL` と target 側の現行 `kernel/KERNEL_VERSION` を比較：

- `REQUIRES_KERNEL > 現行 kernel` の場合、**先に kernel bundle を overlay する** か、当該モジュール更新を block してユーザに kernel 更新を促す
- この順序を守らないと、新モジュールが古いカーネルの未提供 API を呼んで実行時エラーになる

---

## 新モジュール / 欠損モジュールの扱い

| 状況 | 動作 |
|---|---|
| **template にあり target にない** | overlay で追加（`module.csv` / `preset.csv` / すべての `.ps1` / `Guide.txt` / `VERSION` / `REQUIRES_KERNEL`）。`_list.csv` は copied されない（operator が Fabriq Studio で新規作成） |
| **target にあり template にない** | site-custom モジュールと見なし **保持**（touched しない） |

---

## 更新前の安全チェック（外部ツール推奨）

| チェック | 目的 |
|---|---|
| `Fabriq.exe` プロセスが実行中でないこと | ファイルロック回避 |
| `kernel/json/resume_state.json` 不在 | キッティング中断中の更新を避ける |
| target の `kernel/KERNEL_VERSION` と全 module の `REQUIRES_KERNEL` の整合 | 更新後のランタイム互換性保証 |
| target フォルダのバックアップ取得 | ロールバック経路確保 |

---

## schemaVersion の後方互換

`dev/framework_overlay_rules.json` の `schemaVersion` フィールドが将来 `2` 等に上がった場合、外部ツールは未対応バージョンを検知したら **処理を拒否して明示エラー** を返す責任がある（黙って部分動作しない）。

現行 `1` のスキーマは下位互換を維持する形で進化させる方針。

---

## 二つの patch flavor（feedback memory `feedback_patch_creation` 由来）

fabriq には実運用で 2 種類の patch があり、状況で使い分ける：

### 1. Targeted Patch（手動ミラー）

- 単一 / 少数のファイルだけを対象
- `dev/build_framework_patch.ps1` を介さず、手動で配布物を作る
- 例: 単一モジュールの hotfix、ドキュメント修正

### 2. Framework Patch（dev/build_framework_patch.ps1）

- bundle 単位の系統的更新
- `framework_overlay_rules.json` の契約に従って kernel 全体 or モジュール群を 1 zip 化
- fabriq_studio の update 機能から消費される
- SemVer 比較で UPDATE / SKIP を bundle ごとに判定可能

「どちらを使うか」は変更スコープと配備状況で operator が判断する（feedback memory 参照）。

---

## fabriq_studio との接点

fabriq 本体（fabriq）と fabriq_studio（GUI 管理ツール、別プロジェクト、WPF / .NET 8）は以下の契約で疎結合：

| 接点 | fabriq 本体側 | fabriq_studio 側 |
|---|---|---|
| マスターパスフレーズ | `kernel/txt/passphrase_verify.txt` を読み | 生成・更新 |
| ホスト管理 | `kernel/csv/hostlist.csv` を読み | 編集（暗号化／復号付き）|
| モジュール設定 | `_list.csv` を読み | 編集（preset.csv ドロップダウン UI 提供） |
| プロファイル管理 | `profiles/*.csv` を読み | 編集（マーカーパレット UI） |
| 更新オーバーレイ | overlay 対象 | `framework_overlay_rules.json` を読み bundle 比較 → patch 適用 |
| レジストリカタログ | `reg_template` モジュールが消費 | カタログから workspace へエクスポート |

fabriq 本体は Studio のバージョン・機能に依存しない（疎結合の哲学）。


<!-- ============================================================ -->
# === contracts/evidence_manifest_contract.md ===
<!-- ============================================================ -->

# Evidence Manifest 契約（External Evidence Consumer ↔ fabriq）

`KERNEL_API.md §10` および `kernel/EVIDENCE_MANIFEST.md` で公式宣言された、**外部 evidence consumer ツール**（代表: `fabriq_evidence_manager`、別プロジェクト C#/WPF/.NET8）が前方互換に消費するための公開契約。

`evidence_config` モジュール v1.3.0+ が `pc_information/<dir>/manifest.json` を出力。kernel 2.2.2 で contract 公開化。

---

## ファイル配置

```
{evidenceBaseDir}/pc_information/{collectionDir}/manifest.json
```

`{collectionDir}` は `{date}_{uid}_{pc}` 命名（legacy / unified どちらも同じ）。

- 1 evidence_config 実行 = 1 manifest.json
- 再実行時は既存 manifest を `manifest.json.bak` に rename してから上書き（**1 世代のみ保持**、`.bak.bak` は作らない）
- manifest 内のファイルパスは **manifest.json 自身からの相対パス**（self-contained）

---

## schemaVersion=1 のスキーマ

```json
{
  "schemaVersion": 1,
  "manifestType": "fabriq-evidence-manifest",
  "evidenceConfigVersion": "1.3.0",
  "fabriqKernelVersion": "3.2.2",
  "collectedAt": "2026-04-25T13:28:39+09:00",
  "computerName": "NEW-PC-01",
  "hardwareUniqueId": "T2NXCV06Y22208C",
  "selectedNewPcName": "NEW-PC-01",
  "workerName": "suzuki",
  "sections": [
    {
      "id": "01",
      "title": "System Basic Info",
      "files": ["01_SystemInfo.txt"],
      "status": "Success",
      "reason": null,
      "elapsedMs": 145
    },
    {
      "id": "14",
      "title": "Server Roles & Features (CSV)",
      "files": [],
      "status": "Skipped",
      "reason": "Client OS detected (Server-only section)",
      "elapsedMs": 12
    },
    {
      "id": "22",
      "title": "Office License / Activation Status",
      "files": ["22_OfficeLicense.txt"],
      "status": "Success",
      "reason": null,
      "elapsedMs": 8200
    }
  ],
  "summary": {
    "sectionCount": 23,
    "successCount": 21,
    "skippedCount": 2,
    "failedCount": 0,
    "partialCount": 0
  }
}
```

### トップレベルフィールド

| Field | Type | Required | Description |
|---|---|---|---|
| `schemaVersion` | int | yes | 現行 `1`、破壊的変更時に `2` |
| `manifestType` | string | yes | 固定値 `"fabriq-evidence-manifest"`（type discrimination 用） |
| `evidenceConfigVersion` | string | yes | manifest を書いた evidence_config モジュールの SemVer |
| `fabriqKernelVersion` | string | yes | manifest 書き込み時点の `kernel/KERNEL_VERSION` |
| `collectedAt` | string (ISO 8601) | yes | 収集開始日時。タイムゾーンオフセット付き |
| `computerName` | string | yes | `$env:COMPUTERNAME`（実 OS 上のコンピュータ名） |
| `hardwareUniqueId` | string | yes | `Get-HardwareUniqueId` 戻り値（BIOS Serial 由来） |
| `selectedNewPcName` | string | yes | `$env:SELECTED_NEW_PCNAME`（無ければ `computerName` と同値） |
| `workerName` | string \| null | no | `$env:FABRIQ_WORKER_NAME`（profile 実行外では null） |
| `sections` | array<Section> | yes | セクション結果配列 |
| `summary` | Summary | yes | 集計値 |

### Section オブジェクト

| Field | Type | Required | Description |
|---|---|---|---|
| `id` | string | yes | セクション ID。`"01"`〜`"22"`、`"8b"` 等の副 ID も許容 |
| `title` | string | yes | セクション名（例: `"System Basic Info"`） |
| `files` | string[] | yes | manifest 親ディレクトリからの相対パス配列。`/` で終わる文字列はディレクトリを意味する。書き込みファイルが無ければ空配列 |
| `status` | enum | yes | `"Success"` / `"Skipped"` / `"Failed"` / `"Partial"` のいずれか |
| `reason` | string \| null | yes | 非 Success 時の理由文字列。Success 時は `null` |
| `elapsedMs` | int | yes | セクション処理経過時間（ミリ秒） |

### Status セマンティクス

- **Success**: セクション完了、`files[]` のすべてが書き込まれた
- **Skipped**: 意図的にスキップ（例: Server 専用セクションを Client OS で実行 / Office 未インストール / Defender 未稼働）。`reason` 必須、`files` は通常空
- **Failed**: 例外発生でセクションが完了できなかった。`reason` に exception メッセージ。`files` は **常に空配列**（途中まで書かれた壊れたファイルを manifest に載せないため）
- **Partial**: 単一セクション内で複数の独立処理（例: §11 DesktopApps + StoreApps）の一部だけ成功した状態。`reason` に詳細

### Summary オブジェクト

| Field | Type | Required | Description |
|---|---|---|---|
| `sectionCount` | int | yes | `sections.length` |
| `successCount` | int | yes | status=Success の数 |
| `skippedCount` | int | yes | status=Skipped の数 |
| `failedCount` | int | yes | status=Failed の数 |
| `partialCount` | int | yes | status=Partial の数 |

**不変条件**: `successCount + skippedCount + failedCount + partialCount === sectionCount`

---

## 前方互換ルール

### 外部ツール（manager 等）の責任

1. **schemaVersion チェック必須**: 未知 major 版を検知したら警告を出し、legacy mode（manifest 無視 + ファイル列挙）にフォールバックする。silent な部分動作は禁止
2. **未知 section ID は raw 表示**: パーサが知らない `id` のセクションは raw text/CSV としてそのまま提示する。クラッシュさせない
3. **未知 status enum 値は Failed 扱い**: 将来 `"InProgress"` 等が追加されても安全側に倒す
4. **追加フィールドは無視**: schemaVersion=1 内での後方互換な field 追加は manager 側で無視可能

### evidence_config 側の責任

1. **schemaVersion を上げない限り破壊しない**: フィールドの削除・改名・型変更は schemaVersion=2 への昇格を伴う
2. **新 section 追加は schemaVersion=1 内で OK**: 既存 manager は未知 ID として raw 表示するので clash しない
3. **status enum 拡張は schemaVersion=2**: 既存 4 値以外を追加する場合のみ schemaVersion を上げる
4. **任意フィールドの追加は schemaVersion=1 内で OK**: required は変えない

---

## 再実行時の挙動

evidence_config を同じディレクトリで再実行する際:

1. 既存 `manifest.json` が存在すれば `manifest.json.bak` に rename（既存 `.bak` は削除して上書き）
2. 新しい収集を実行し、新 manifest を atomic に書き出し（収集完了時に一括）
3. 中断時の半端な manifest を防ぐため、incremental write は採用しない

---

## ディレクトリ表現

`files[]` の要素が `/` で終わる場合、それはディレクトリを意味する：

```json
"files": ["20_TempBackup.txt", "20_TempBackup/"]
```

manager は `/` で終わる要素を「opaque な forensic dump dir」として扱い、内部のファイルは個別パースしない（必要なら raw として一覧化のみ）。

---

## manifest 不在の旧 evidence

manifest.json 不在の旧形式 evidence（kernel 2.2.1 以前 / evidence_config 1.2.0 以前で収集されたもの）も外部ツールはサポートし続けることが期待される。manifest が無ければファイル列挙ベースで動作する従来挙動を維持する。

---

## 監査ギャップとの関連

`evidence_config` v1.2.0 は 22 セクションを出力するが、官公庁向け監査では追加項目が必要（project memory `project_evidence_audit_gaps` 参照）：

- **§23–§30 候補**: TPM / SecureBoot / gpresult / secedit / 証明書 / NTP / AccessControl

これらの追加は schemaVersion=1 内で OK（後方互換 section 追加）。manifest 契約は将来 audit gap を埋める追加に耐えられる設計になっている。

evidence_manager 側の対応では、§21（License）/§22（Office License）の重要 section 表示が manifest schema 経由で必須化されることで、過去の "files only" 表示モードでの section 取りこぼしを防ぐ（project memory `project_evidence_manifest_gap` 参照）。


<!-- ============================================================ -->
# === apps/00_apps_overview.md ===
<!-- ============================================================ -->

# fabriq apps カタログ

`e:/fabriq/apps/` 配下には、fabriq カーネルおよびモジュール群とは独立した GUI サブプロジェクト群が並んでいます。これらは「補助具 (auxiliary tools)」であり、いずれも単体の `.ps1` をエントリポイントとして起動できる自己完結型の WinForms アプリです。

apps の特徴:

- **fabriq_operator のみがダッシュボード本体**で、`kernel/main.ps1` から dot-source されて起動される。それ以外はすべて操作員が任意のタイミングで起動する独立アプリ。
- **配布**: フレームワーク本体に同梱され、`apps/` ディレクトリごと運搬される。
- **発見**: `fabriq_operator` の Settings タブから [And More...] → [FabriqApps] を開くと、`apps/<name>/<name>.ps1` 規約で並んだ各アプリが自動的にリストされる (`apps_dialog.ps1`)。`fabriq_operator` 自身と `fabriq_ios` は除外される。
- **テーマ**: 大半は CentreCOM 風ライトテーマまたは Fabriq 標準ダークテーマで作られており、操作員の視覚的同一性が保たれる。
- **ファイル契約**: 各 GUI は対応する CSV (例: `bloatware_list.csv`, `local_user_list.csv`) を直接編集する。スタジオを介さずに現場で CSV 編集を完結させるための「現場用ツール」。

以下、各アプリの目的と存在理由を列挙します。

## fabriq_operator

`apps/fabriq_operator/` — **fabriq 全体のメインダッシュボード**。`fabriq_operator.ps1` 自身は薄いブートストラップで、`lib/` 以下に分割された 6 ファイルを dot-source する構成 (theme / session_form / apps_dialog / quickactions_dialog / dashboard_form / flex_dashboard)。`kernel/main.ps1` がこのファイルを読み込んだ瞬間に WinForms アセンブリが投入され、`Show-SessionSetupForm` および `Show-OperatorDashboard` が公開される。fabriq の操作員がキッティング作業中にもっとも多く触るアプリで、Profile 実行 (Linear / Flex)、モジュール単発実行、Quick Actions、CSV Editor / Windows Update / Refabriq などの dispatch をすべて担う。詳細は `01_fabriq_operator_dashboard.md` を参照。

## fabriq_ios

`apps/fabriq_ios/` — **Cisco IOS 風コマンドラインシェル**。fabriq の上に被せた「シュルキティニスム宣言の芸術部門」と SPEC.md にある通り、実用ツールではなく art object としての位置付け。User EXEC → Privileged EXEC → Global Config → Interface / Module Config の 4-5 階層モード遷移、タブ補完 (PSReadLine 統合)、Cisco 風 syslog（メッセージはシュルレアリスム的な英文）を備える。`fabriq_ios.ps1` は self-spawn ガード付きで、`Start-Process powershell.exe` で隔離サブプロセスを立ち上げ、PSReadLine の KeyHandler や環境変数が呼び出し元に漏れないようにしている。fabriq_ios のみ独自の `VERSION` (現在 0.3.5) を持ち、kernel SemVer とは独立して進化する。詳細は `02_fabriq_ios.md`。

## csv_editor

`apps/csv_editor/csv_editor.ps1` — **ジェネリック CSV エディタ**。fabriq_studio が無い現場での編集経路。20 種類以上の編集対象 CSV (hostlist.csv, app_list.csv, reg_hkXm_list.csv, local_user_list.csv 等) を `$script:CsvRegistry` に静的登録しつつ、`profiles/*.csv` も動的発見する。ListView でファイル選択 → DataGridView で内容編集 → Save の 3 ペイン構成。

## system_launcher

`apps/system_launcher/system_launcher.ps1` — **Windows 設定ショートカットのパレット**。`ms-settings:` URI、`*.cpl`、`*.msc`、shell:: GUID、cmd / powershell など 34 項目の Windows システムツールを 1 画面に集約し、検索ボックスでフィルタしてダブルクリック起動できる。Windows Search を使わずに辿り着くことで「Search 履歴を残さない」のが運用上のメリット (キッティング後の証跡をクリーンに保つ意図)。

## bloatware_exporter

`apps/bloatware_exporter/bloatware_exporter.ps1` — **インストール済み Win32 アプリのスキャン & エクスポート GUI**。HKLM/HKCU の `Uninstall` レジストリキーを走査して、`bloatware_remove` モジュールが食う `bloatware_list.csv` を編集できるようにする。スキャン結果から不要アプリを ✓ するだけで除去対象 CSV が育つ。

## desktop_icon_backup_app

`apps/desktop_icon_backup_app/desktop_icon_backup_app.ps1` — **デスクトップアイコン配置のスタンドアロン バックアップツール**。`HKCU\Software\Microsoft\Windows\Shell\Bags\1\Desktop` を `.reg` 形式でエクスポートし、`modules\extended\desktop_icon_config\backup\` に保存する。`desktop_icon_config` モジュールが復元側を担うので、このアプリは「採取側」を独立 GUI で提供する位置付け。

## local_user_setup

`apps/local_user_setup/local_user_setup.ps1` — **ローカルユーザー作成ウィザード**。`local_user_list.csv` 内のプレースホルダー行 (UserName / Password が空の行) を 1 件ずつ順に埋めていく Wizard 型 GUI。明るいライトテーマで、配布前の CSV プリセット段階で使うことを想定。

## storeapp_editor

`apps/storeapp_editor/storeapp_editor.ps1` — **Store アプリ削除リストエディタ**。`Get-AppxPackage` 系で取得したインストール済み Store アプリ一覧と `storeapp_list.csv` を並べ、削除対象の追加・削除・編集を行う。`storeapp_config` モジュールの入力ファイルメンテナンスが目的。

## winget_gui

`apps/winget_gui/winget_gui.ps1` — **winget パッケージ検索 / app_list.csv エディタ**。`winget search` を非同期 Runspace で叩き、ヒットしたパッケージを GUI で確認しつつ `app_list.csv`（`winget_install` モジュールの入力）を編集する。Runspace 利用は GUI のフリーズ回避目的。


<!-- ============================================================ -->
# === apps/01_fabriq_operator_dashboard.md ===
<!-- ============================================================ -->

# fabriq_operator ダッシュボード仕様

`apps/fabriq_operator/` は fabriq の操作員向けメインダッシュボードです。`fabriq_operator.ps1` は WinForms アセンブリを読み込み、`lib/` 以下の 6 ファイルを dot-source するだけのブートストラップで、機能本体はすべて `lib/*.ps1` 側に分散しています。

## ファイル構成

| ファイル | 役割 |
|---|---|
| `fabriq_operator.ps1` | WinForms 初期化 + lib/*.ps1 のロード |
| `lib/theme.ps1` | カラーパレット / フォント / `New-Styled*` ヘルパー群 |
| `lib/session_form.ps1` | セッション開始フォーム (`Show-SessionSetupForm`) |
| `lib/dashboard_form.ps1` | メインダッシュボード (`Show-OperatorDashboard`) |
| `lib/flex_dashboard.ps1` | FlexProfile ダッシュボード (`Show-FlexDashboard`) |
| `lib/quickactions_dialog.ps1` | "And More..." サブダイアログ (`Show-AndMoreDialog`) |
| `lib/apps_dialog.ps1` | FabriqApps ランチャー (`Show-AppsDialog`) |

## セッション開始フォーム (`Show-SessionSetupForm`)

ダッシュボードに入る前に必ず表示される。返り値は `@{ WorkerID; WorkerName; SelectedHost; MasterPassphrase; Cancelled }`。

入力項目:

- **Worker** — DataGridView (ID / Name 列、ヘッダクリックで sort 可) で `kernel/csv/worker_list.csv` から選択。一覧外の作業者向けにマニュアル入力テキストボックスも併設し、両者は排他 (片方を入力すると他方をクリア)。
- **Target Host** — DataGridView (AdminID / OldPC / NewPC / IP / Pin)。ライブ検索 (AdminID と NewPCName のみが検索対象。OldPC/IP/Pin は意図的に検索から除外) と件数ラベル "X / Y"。`$env:COMPUTERNAME` と一致する NewPC があれば自動選択し、緑色の「* Auto-detected」ヒントを表示。Row.Tag に host 元オブジェクトを格納することで、sort や filter で行が並び替わっても選択を保持する。
- **Master Passphrase** — `UseSystemPasswordChar` のテキストボックス。`Test-MasterPassphrase` で検証し、失敗時は SelectAll + Focus で再入力を促す。

キーボードショートカット (この form だけは検索効率のためマウスオンリーから外れている):

- 検索 BOX で Esc → 検索クリア
- 検索 BOX で Enter → passphrase BOX へフォーカス移動
- passphrase BOX で Enter → [Start Session] 押下

## メインダッシュボード (`Show-OperatorDashboard`)

`fabriq operator` ウィンドウ (700×560) は **Header + TabControl + StatusBar** の 3 段構成。Header には CentreCOM 風青/黄/赤の 3 色アクセントストライプが入り、右側に `HostName | W: WorkerName` が表示される。

返り値の Action 値: `Quit / ExecuteProfile / FlexProfile / ExecuteModules / NewSession / OpenCsvEditor / WindowsUpdate / Restart / Refabriq / HistoryExport / RegenerateChecklist / Manifesto / SystemLauncher / AppsMode`。これらは `kernel/main.ps1` の switch でディスパッチされる。

### Tab 1: Profiles

| 要素 | 役割 |
|---|---|
| Profile DataGridView | ProfileName / Modules / Total / Path (hidden) の 4 列。`Load-Profiles` で `profiles/*.csv` を走査して充填 |
| Profile Detail RichTextBox | 選択行の中身を `Order Description` 形式で表示 (リアルタイム) |
| `[AutoPilot]` チェックボックス | デフォルト ON。Linear 実行時にのみ参照される |
| `[View Details]` | 選択 Profile の CSV を Explorer で `select` 表示 |
| `[Execute (Flex)]` | 緑ボタン。Action=`FlexProfile` を返す。AutoPilot は伝播しない (Flex 側で独立管理) |
| `[Execute Profile]` | 青ボタン (Linear path)。Action=`ExecuteProfile`。AutoPilot チェックボックス値を伝播 |

### Tab 2: Modules

| 要素 | 役割 |
|---|---|
| Category ComboBox | "All" + `$GroupedModules` の name を列挙 |
| Search TextBox | MenuName 部分一致 |
| Module DataGridView | # / Module / Category 列 + Script / Order 列 (hidden)。`Update-ModuleGrid` で再構築 |
| `[Execute]` | 単独行を選択して Action=`ExecuteModules`、`SelectedModules=@($script:allModuleData[$idx])` を返す |
| 行ダブルクリック | 同上 (1 クリック相当) |

### Tab 3: Settings

縦に並んだセクション構成:

- **Evidence Output Path** — `$global:FabriqEvidenceBasePath` を表示。`[Open Folder]` で Explorer 起動 (path 未生成時は親ディレクトリへフォールバック)。
- **Quick Actions (フロント行 5 + And More)**:
  - `[Open CSV Editor]` → Action=`OpenCsvEditor`
  - `[Windows Update]` → Action=`WindowsUpdate`
  - `[Refabriq]` → Action=`Refabriq`
  - `[System Launcher]` → Action=`SystemLauncher`
  - `[And More...]` → `Show-AndMoreDialog` をモーダル表示。返り値の Action (Restart / HistoryExport / RegenerateChecklist / AppsMode) を $result に転記
- **Session** — `Worker / Host / KanriNo` 表示 + `[New Session]` (Action=`NewSession`)
- **Manifesto** — `[Manifeste du Surkitinisme]` (Action=`Manifesto`)

### Status Bar

`$LastResultSummary` を表示するだけの最下段ストリップ。空のときは "Ready"。

## FlexProfile ダッシュボード (`Show-FlexDashboard`, kernel 3.1.5+)

Linear 経路の単方向実行に対し、**1 モジュール単位で再実行可能な状態管理 GUI**。`Resolve-ProfileModules ... -IncludeDisabled` で Profile 全行を取り、history.csv と最後の `$LastBatchResults` を上書きで反映した stateMap (Order → Status/Verified/Message) を作る。

返り値の Action: `Close / RunSingle / RunBatch / RunGroup / Complete / RestartNow / ResetState`。

### グリッド構成 (左→右)

| 列 | 役割 |
|---|---|
| Checked | 選択チェックボックス。デフォルトすべて未チェック (Flex 哲学: 白紙からピック) |
| # | Order |
| Group | 3.2.0+ の Group バー連動 (空欄ありのケースは `_Group` 列が無い旧 CSV) |
| Module | MenuName |
| Status | Pending / Success / Partial / Error / Skipped / Cancelled。CellFormatting で badge 描画 (Success=緑, Partial=黄, Error=赤, Skipped/Cancelled=グレー, Pending=薄グレー) |
| Verified | -/PASS/FAIL。PASS=緑バッジ, FAIL=赤バッジ |
| Run | 行内ボタン。クリックで RunSingle dispatch (3.1.8 で footer の "Run This: M" を置き換え) |

右クリックメニュー: 「Mark as Pending (reset state)」が 1 項目。Action=`ResetState` で発火。

### Groups バー (3.2.0+)

Profile CSV に `Group` 列があり値が入っている場合、Header 直下に `[Run: <GroupName>]` ボタンを CSV 出現順で並べる。クリックで全 Group メンバーの Order を集めて `Action=RunGroup`、AutoPilot=true 強制で発火。横幅は 130px × N、最大 ~6 個まで panel に収まる (auto-wrap 無し、YAGNI 原則)。

### Footer

| 行 | 要素 |
|---|---|
| 上段 | `[Select All]` (Enabled=1 行のみ ✓) / `[Clear All]` / WaitSec NumericUpDown / `[Run Selected (N)]` 大ボタン / `[Restart Now]` |
| 下段 | `[Complete]` 大ボタン / `[Back]` |

`[Complete]` の文言は status に応じて動的変化:

- Issue (Error / Partial / 未実行で ✓ された Pending) があれば → 黄色 + "Complete with N issues"
- 何も実行していない (全 Pending) → 黄色 + "Complete (nothing executed)"
- それ以外 → 緑 + "Complete"

`[Complete]` 押下時、issue があれば Yes/No 確認ダイアログを出し、Yes ならフィナライズ進行。

### PendingFinalize 警告

`-PendingFinalize $true` で開かれた Flex は、Header の "Last finalized:" の代わりに **赤バッジ "PENDING FINALIZE"** を表示。X/Back で閉じようとすると「checklist が生成されない / evidence が upload されない」警告ダイアログを出して操作員に Complete を促す。

## "And More..." ダイアログ (`Show-AndMoreDialog`)

420×320 のサブダイアログ。Restart PC / Export History / Regenerate Checklist / FabriqApps の 4 項目を DataGridView で表示し、選択 → `[Launch]` で Action 文字列を main.ps1 と同じ語彙で返す。アクションは Cancel / Restart / HistoryExport / RegenerateChecklist / AppsMode。

## FabriqApps ダイアログ (`Show-AppsDialog`)

`apps/<name>/<name>.ps1` 規約で `apps/` 配下の各サブディレクトリを走査し、`fabriq_operator` と `fabriq_ios` を除外した一覧を提示する。`[Launch]` で Action=`Launch`, AppName, AppPath を返し、main.ps1 が `& $appPath` で起動する。

## テーマシステム (`theme.ps1`)

CentreCOM 風ライトテーマ。fabriq_evidence_manager と同じデザイン言語で揃えてある。

### カラーパレット (主要)

| 名前 | 値 (ARGB) | 用途 |
|---|---|---|
| `bgForm` | #BABEC2 | フォーム背景 |
| `bgPanel` | #4A4A4A | ヘッダーバー |
| `bgGrid` | #E1E4E7 | DataGridView 背景 |
| `bgCellAlt` | #EDEEEF | 交互行 |
| `bgHeader` | #A0A6AB | グリッドヘッダー |
| `bgButton` | #989DA1 / `bgButtonHov` #AAAEB3 | 通常ボタン |
| `bgAccent` | #4A90D9 | アクセント青 (Execute Profile 等) |
| `bgAdd` | #4CAF50 | 成功緑 (Execute Flex / Complete) |
| `bgDelete` | #C62828 | エラー赤 |
| `bgSelection` | #4CAF50 | 行選択時 |
| `bgTabPage` | #C4C8CC | タブ本体 |
| `stripeBlue/Yellow/Red` | #4A90D9 / #F2C94C / #EB5757 | ヘッダー 3 色ストライプ |

### フォント

- `fontNormal` Segoe UI 9pt
- `fontBold` Segoe UI 9pt Bold
- `fontSemiBold` Segoe UI Semibold 9pt
- `fontLarge` Segoe UI 11pt Bold
- `fontTitle` Segoe UI 14pt Bold
- `fontMono` Consolas 8.5pt

### 提供ヘルパー

- `New-StyledButton` — 角無し Flat、border = `borderColor`、accent 色のときは白文字
- `Set-GridStyle` — DoubleBuffered ON、SelectionMode=FullRowSelect、MultiSelect=false、ReadOnly=true
- `Set-TextBoxStyle`, `New-StyledLabel`, `New-StyledPanel`, `New-StyledCheckBox`, `New-StyledComboBox`, `Set-FormStyle`

これらヘルパーが Operator GUI の見た目を一手に管理している。CentreCOM 風 (キャリアグレード機器ライク) の青/黄/赤ストライプは fabriq の視覚的アイデンティティとして session_form / dashboard_form の両方に踏襲される。


<!-- ============================================================ -->
# === apps/02_fabriq_ios.md ===
<!-- ============================================================ -->

# fabriq_ios — Cisco IOS 風シェル

`apps/fabriq_ios/` は fabriq フレームワーク上に被せた **Cisco IOS スタイルのコマンドラインシェル**です。SPEC.md にて「シュルキティニスム宣言の芸術部門」と明確に位置付けられており、実用ツールではなく **art object / 思想の戯画化コンポーネント**として存在します。それでもタブ補完や省略補完は本物の Cisco IOS に近い水準で作り込まれています。

## VERSION (独立 SemVer)

`apps/fabriq_ios/VERSION` は現在 **0.3.5**。fabriq カーネル / モジュールの SemVer とは独立して進化します。`show version` および `show running-config` コマンドが参照するほか、KERNEL_API.md §6 の internal API (`Invoke-BatchExecution`, `Initialize-ModuleSystem`, `Resolve-ProfileModules`, `Test-MasterPassphrase`) に依存しているため、kernel PATCH 昇格のたびに再検証する旨が README に明記されています。

## エントリポイントと self-spawn 隔離

`fabriq_ios.ps1` の冒頭に **self-spawn ガード**があります。

```
if (-not $env:FABRIQ_IOS_SUBPROCESS) {
    $env:FABRIQ_IOS_SUBPROCESS = '1'
    Start-Process powershell.exe -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$self -Wait
    return
}
```

これは fabriq_operator の FabriqApps ダイアログから `& $appPath` で in-process 起動された場合に、**PSReadLine の KeyHandler / 環境変数の改変 / global スコープ変数**が呼び出し元の powershell に漏れないようにするためです。`-NoProfile -ExecutionPolicy Bypass -File` で隔離された子 powershell.exe がシェルを実行し、終了すると環境変数も連れて消えます。

その後、kernel/common.ps1 を dot-source してから次の関数を **no-op で global シャドウ**します:

- `Initialize-ExecutionHistory`, `Restore-ExecutionHistory`, `Write-ExecutionHistory`, `Add-ExecutionResult`
- `Export-ExecutionHistory`, `Export-HtmlChecklist`
- `Initialize-EvidenceBasePath`, `Capture-ScreenEvidence`

これにより、モジュールが間違って呼んでも fabriq_ios の操作が親 fabriq の audit trail を汚さない保証が立ちます (REPL は履歴を残さない)。

## モード構造 (5 階層)

`lib/shell_state.ps1` の `New-ShellState` で初期 Mode='UserExec'。許可遷移は `Set-ShellMode` の whitelist。

```
UserExec        fabriq>            参照系のみ。enable で昇格
PrivilegedExec  fabriq#            disable / configure terminal / reload / show
GlobalConfig    fabriq(config)#    hostname / interface / module / cleanup / copy / install / script
InterfaceConfig fabriq(config-if)# ip address ...
ModuleConfig    fabriq(config-mod)# / (config-clean)# / (config-copy)# / (config-install)# / (config-script)#
```

ModuleConfig の prompt suffix は `data/module_categories.json` の各 category の `promptSuffix` を引いて動的決定 (例: `module bitlocker_config` で入った場合は config-mod、`cleanup directory_cleaner` なら config-clean)。

各モードのディスパッチャは `lib/modes/<mode>.ps1` に分離 (`Invoke-UserExecCommand` 等)。

## コマンドパーサ (`lib/parser.ps1`)

Cisco IOS の prefix-resolution を厳密に再現:

- `ConvertTo-FabriqIosTokens` — quote 対応の単純なトークナイザ
- `Get-FabriqIosCommandVocabulary` — モードごとのコマンド一覧 (UserExec=`enable show help exit ?`, PrivilegedExec=`show configure reload disable exit help ?` など)
- `Resolve-Token` — 完全一致を優先し、prefix 一致が 1 候補のときだけ採用、複数候補は ambiguous エラー (`% Ambiguous command`)
- `Get-FabriqIosSubVocabulary` — 2 層目語彙 (`PrivilegedExec.show` → `version running-config profiles modules evidence manifesto` など)

これにより `conf t`、`sh ru`、`int eth0` のような Cisco 流の省略入力が動きます。

## タブ補完 (`lib/completer.ps1`)

PSReadLine 統合。`Register-FabriqIosCompleter` がシェル開始時に Tab と `?` をフックします。コア関数 `Get-FabriqIosCompletion` は **pure function** で、(Line, Position, Mode, State) を取り候補配列を返すだけ。PSReadLine 抜きで単体テスト可能 (実際 `tests/completer.tests.ps1` がそのテストハーネス)。

動的補完ソース:

- `hostname <TAB>` → `kernel/csv/hostlist.csv` の NewName 列
- `interface <TAB>` → `Get-NetAdapter` の InterfaceAlias (日本語「イーサネット」も候補に出る — むしろ味)
- `show <TAB>` → モード別 sub-vocabulary
- `module <TAB>` / `cleanup <TAB>` / `install <TAB>` 等 → `module_categories.json` の対応 category 内のモジュール名
- ModuleConfig の `set <col> <val>` / `add <col> <val>` 系は **位置パリティ**で次候補を col 名 / 値 enum に切り替え

`?` キーは Cisco 慣習を再現:

- 空 prompt で `?` → そのモードの全 vocabulary を help_text.csv から引いて表示 (deferred AcceptLine 経由)
- 入力途中で `?` → 候補一覧を表示してバッファを復元、編集を続行可能

## 主要コマンド (`lib/commands/*.ps1`)

| ファイル | 提供コマンド |
|---|---|
| `show.ps1` | show version / show host(s) / show running-config / show profiles / show modules / show evidence / show manifesto |
| `enable_disable.ps1` | enable (passphrase 認証で PrivilegedExec へ) / disable (UserExec へ降格) |
| `hostname.ps1` | hostname `<NewName>` で hostlist エントリ選択 + hostname_config 起動 |
| `interface.ps1` | interface `<alias>` で InterfaceConfig 入り |
| `ip_address.ps1` | ip address from-hostlist / ip address `<ip> <mask>` |
| `categories.ps1` | module / cleanup / copy / install / script の category 切替 |
| `module.ps1` | module `<name>` で ModuleConfig 入り |

`enable` は `Test-MasterPassphrase` を流用するため、fabriq の DPAPI 暗号化済みパスフレーズと同一のものを要求します。

## syslog システム (`lib/syslog.ps1` + `data/syslog_messages.csv`)

Cisco IOS 完全準拠の syslog 行を生成します:

```
*Apr 29 14:23:01.234: %FABRIQ-5-HOSTNAME: The name NEW-PC-01 has been carved into the silicon. The machine awaits the next dawn.
```

書式:

- Timestamp は `*MMM dd HH:mm:ss.fff` (英語 month 強制、locale 非依存)
- Severity 0–7 (Cisco 準拠: Emergency / Alert / Critical / Error / Warning / Notification / Informational / Debug)
- Mnemonic = `FABRIQ-X-<FACILITY>` (HOSTNAME, INTERFACE, IPADDR, MANIFESTO, AUTOMATE, RESTART, ENABLE, EXIT)

メッセージテンプレートは `data/syslog_messages.csv` で外部編集可能。例:

```
HOSTNAME,5,success,The name {NewName} has been carved into the silicon. The machine awaits the next dawn.
HOSTNAME,4,reboot_required,The machine must close its eyes and open them again to remember its new name.
IPADDR,6,whispered,Address {Ip}/{Prefix} has been whispered to interface {Interface}. The wire understands.
INTERFACE,5,configured,"Interface {Alias}, the metallic vein, is now configured."
MANIFESTO,7,quote,"Surkittinism is the convulsive beauty of mass deployment, or it is nothing."
```

設計思想 3 本柱 (SPEC.md より):

1. **タブ補完は真面目に** — 操作体験は本物の Cisco IOS 水準
2. **syslog メッセージは Cisco 風英語、ただしシュルレアリスム** — 形式厳格・内容詩的
3. **コマンド階層は Cisco IOS 丸パクリ**

## バナー (`data/version_banner.txt`)

起動時に表示される:

```
Fabriq IOS Software, Version 3.0(1)Surkittinism
Copyright (c) 1924-2026 by Andre Breton & Anonymous Kitting Operators.
Compiled in the Dream Hours by automatic-writing.

The machine yawns. Press RETURN to begin the seance.
```

著作権年 1924-2026 は André Breton の『シュルレアリスム宣言』(1924) を起点としており、fabriq_ios の "art object" としての自意識を体現しています。

## テスト (`tests/`)

Pester 5 ベース。モジュール完成度ごとに `_phase3_smoke.ps1` ... `_phase9b_smoke.ps1` および `parser.tests.ps1`, `completer.tests.ps1`, `prompt.tests.ps1`, `shell_state.tests.ps1`, `syslog.tests.ps1` が存在。`Get-FabriqIosCompletion` 等の pure function はそのまま単体テスト可能。

## 存在理由

「真面目なフレームワークの中に冗談を一つ混ぜる」「Cisco ルータと日本語 Windows キッティング現場という本来交わらない二つの世界をロートレアモン的に手術台の上で出会わせる」(SPEC.md より要約)。kernel 内部 API への結合は「the price of the joke」として README 上で意識的に受容されており、kernel PATCH のたびに再検証することが運用上のルール。

```
fabriq>                              User EXEC mode (read-only)
fabriq#                              Privileged EXEC mode (after enable)
fabriq(config)#                      Global Configuration mode (after configure terminal)
fabriq(config-if)#                   Interface Configuration mode (after interface XXX)
fabriq(config-mod)#                  Module Configuration mode (after module XXX)
```


<!-- ============================================================ -->
# === apps/03_other_apps.md ===
<!-- ============================================================ -->

# その他の補助 GUI アプリ

`fabriq_operator` と `fabriq_ios` を除く 7 つの apps は、いずれも **fabriq_operator の Settings タブ → [And More...] → [FabriqApps]** 経由で起動する独立 WinForms アプリです。直接 `apps/<name>/<name>.ps1` を実行することもでき、起動経路は対称的です。

それぞれが特定の CSV (kernel/csv/ または modules/{standard,extended}/<module>/*_list.csv) を編集対象として固定的に持っており、「Studio が現場に無い場合の代替経路」または「専門用途の編集 GUI」を構成します。

## csv_editor

### Purpose
fabriq の主要 CSV 群を一画面で開閉できるジェネリックエディタ。ListView でファイル選択 → DataGridView で内容編集 → Save の 3 段構成で、CSV ごとに異なるカラムスキーマも同じ DataGridView 体験で扱える。

### Trigger
FabriqApps ダイアログから起動するか、ダッシュボード Settings タブの `[Open CSV Editor]` ボタンから直接起動 (Action=`OpenCsvEditor`)。

### Key UI / 機能
- 左ペイン: ListView。Group 列 (Kernel / Apps / Registry / Users / Power / Network / Evidence / Profiles ...) でグループ表示
- 右ペイン: DataGridView。BOM 検出 (`Detect-CsvEncoding`) で UTF-8 / Default を自動判定して読み込み、保存時も同じエンコーディングを維持
- 静的レジストリ `$script:CsvRegistry` に 20 件以上の編集対象を登録 (hostlist.csv, app_list.csv, reg_hkXm_list.csv 各種, local_user_list.csv, power_list.csv, storeapp_list.csv, domain.csv, gyotaku/task_list.csv, ipv6_list.csv, group_list.csv, display_list.csv, dpi_list.csv, license_key.csv, autokey/recipe.csv, copy_list.csv, reg_template/reg_list.csv 等)
- 動的レジストリ: `profiles/*.csv` を起動時にスキャンして「Profile: <name>」として追加

### CSV touched
fabriq 配下のほぼすべての主要 CSV (上記リスト)。

---

## system_launcher

### Purpose
Windows 設定 / コントロールパネル / 管理ツール / シェルへのショートカットを 1 画面に集約したパレット。**Windows Search を使わずに辿り着く**ことで、検索履歴をキッティング先 PC に残さない運用上のメリットがある (apps と commands 両方に同じスクリプトが存在し、commands 側は Status Monitor の手動操作からも呼べる)。

### Trigger
ダッシュボード Settings の `[System Launcher]` (Action=`SystemLauncher`) または FabriqApps ダイアログ。

### Key UI / 機能
- 検索ボックス + DataGridView の 2 ペイン (Num / Name / Category)
- 34 項目を Settings / Control Panel / System Tools / Shell の 4 カテゴリで保持
- `ms-settings:`, `*.cpl`, `*.msc`, `shell:::{GUID}` (God Mode), `cmd.exe`, `powershell.exe`, `runas` 起動 を Type 列で切替
- `Invoke-Tool` 関数が Type に応じて Start-Process の起動形態を切替 (uri / shell / runas / exe)

### CSV touched
なし。

---

## bloatware_exporter

### Purpose
インストール済み Win32 アプリ (HKLM/HKCU の Uninstall レジストリキー) を GUI でスキャンし、`bloatware_remove` モジュールが食う `bloatware_list.csv` を編集するエディタ。

### Trigger
FabriqApps ダイアログ。

### Key UI / 機能
- スキャン側 DataGridView (現在インストール済みアプリ一覧) と CSV 側 DataGridView (削除候補リスト) の 2 ペイン構成
- `[Add to CSV]` で行を移送、`[Remove]` で CSV から取り除く
- `$script:csvPath` は固定で `..\..\modules\standard\bloatware_remove\bloatware_list.csv` を絶対パス解決
- `$script:isDirty` フラグで未保存変更検出

### CSV touched
`modules/standard/bloatware_remove/bloatware_list.csv`

---

## desktop_icon_backup_app

### Purpose
デスクトップアイコン配置を保持するレジストリキーを `.reg` 形式でエクスポートするスタンドアロン GUI。`desktop_icon_config` モジュールが復元側を担うので、このアプリは **採取専用**。

### Trigger
FabriqApps ダイアログ、または **顧客側でこのアプリ単体を渡して採取してもらう**ユースケース (fabriq 本体無しでの利用も意図されている)。

### Key UI / 機能
- 対象キー: `HKCU\Software\Microsoft\Windows\Shell\Bags\1\Desktop`
- バックアップ先: `modules\extended\desktop_icon_config\backup\` (絶対パス解決)
- 大きな `[Backup Now]` ボタン中心の単機能 UI
- 過去バックアップの一覧表示 + Explorer 起動

### CSV touched
なし (`.reg` ファイル出力のみ)。

---

## local_user_setup

### Purpose
`local_user_list.csv` の **プレースホルダー行** (UserName / Password が空欄の予約行) を 1 件ずつ埋めていく Wizard 型 GUI。配布前の CSV プリセット段階での使用を想定し、明るいライトテーマで操作の安心感を強調。

### Trigger
FabriqApps ダイアログ。

### Key UI / 機能
- `Next` / `Back` で 1 アカウントずつ進む Wizard 形式
- UserName / Password / FullName / Description / Group の 5 入力 + 確認用 Re-enter password
- `$script:placeholders` 配列に空行のみ抜き出して進行管理
- 進捗表示「3 / 7」のような counter
- 入力検証: 空文字 / Re-enter mismatch をエラー赤で即時表示

### CSV touched
`modules/standard/local_user_config/local_user_list.csv`

---

## storeapp_editor

### Purpose
`Get-AppxPackage` 系で取得した Store アプリ (`Microsoft.*` / `MicrosoftWindows.*` / OEM AppX 等) を一覧して、`storeapp_config` モジュールの入力 `storeapp_list.csv` を編集する。

### Trigger
FabriqApps ダイアログ。

### Key UI / 機能
- 左: インストール済み AppxPackage 一覧 DataGridView
- 右: 削除リスト CSV エディタ DataGridView
- `[Add to Remove List]` / `[Remove from List]` で双方向移送
- `$script:isDirty` 管理 + 終了時保存確認

### CSV touched
`modules/standard/storeapp_config/storeapp_list.csv`

---

## winget_gui

### Purpose
`winget search <keyword>` を非同期 (Runspace) で叩いてヒットしたパッケージを GUI で確認しつつ、`winget_install` モジュールの入力 `app_list.csv` を編集する。

### Trigger
FabriqApps ダイアログ。

### Key UI / 機能
- 検索ボックス + 検索結果 DataGridView (Id / Name / Version / Source 列)
- 編集対象 CSV DataGridView
- **Runspace 非同期検索** (`$script:runspace` / `$script:asyncHandle`) によって winget 実行中も GUI がフリーズしない
- `[Add to CSV]` で行移送、`[Test winget]` で疎通確認

### CSV touched
`modules/standard/winget_install/app_list.csv`


<!-- ============================================================ -->
# === apps/04_dev_template_and_tooling.md ===
<!-- ============================================================ -->

# dev/ — テンプレート & 開発ツーリング

`e:/fabriq/dev/` は fabriq の **開発支援 / メンテナンス / 配布補助**を担うディレクトリです。実行時には呼ばれず、開発者 (= Claude を含む) が手動で起動するか、release 手順の一部として使われます。

## 構成

```
dev/
├── template/                       新規モジュール作成用スケルトン
│   ├── _template_script.ps1
│   ├── _template_list.csv
│   ├── module.csv
│   ├── Guide.txt
│   ├── VERSION                     0.1.0 (=「開発中・未リリース」の目印)
│   └── REQUIRES_KERNEL             2.0.0 (現行 baseline)
├── framework_overlay_rules.json    contracts に詳述。配布 / 上書きの単一ソース
├── build_framework_patch.ps1       framework patch 生成
├── seed_module_versions.ps1        全モジュールに VERSION/REQUIRES_KERNEL を baseline 打刻
├── check_version.ps1               kernel 版表記の整合性検証
├── verify_comments_only.ps1        コメントのみ変更の証明
├── launcher/                       Fabriq.exe / Fabriq_IOS.exe の C# ソース
│   ├── Launcher.cs / Launcher_IOS.cs
│   ├── app.manifest / app_ios.manifest
│   ├── build.ps1 / build_ios.ps1
│   ├── fabriq.ico
│   └── README.md
└── ico/                            アイコン素材
    ├── app_icon.ico
    ├── app_icon_preview.png
    └── jpg 素材
```

---

## template/ — 新規モジュールスケルトン

CLAUDE.md 絶対遵守事項 #1「標準テンプレートの厳守」の根拠。新規モジュールはここをコピー → `modules/{standard,extended}/<name>/` に配置 → リネームして使う。

### `_template_script.ps1` (7-step 構造)

| Step | 内容 | 必須 |
|------|------|------|
| Step 1 | **CSV 読み込み** — `Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @(...)` で取得。`$null` なら Error、`Count -eq 0` なら Skipped を即返却 | 必須 |
| Step 2 | **前提条件チェック (early return)** — 必要なディレクトリや実行ファイルの存在検証。不要なら丸ごと削除可 | 任意 |
| Step 3 | **Dry-run summary** — 適用対象を [APPLY] / [SKIP] / [NOT FOUND] ラベル付きで列挙。WHAT / EXPECTED OUTCOME を一目で見せる | 必須 |
| Step 4 | **ユーザー確認** — `Confirm-ModuleExecution -Message "..."` で Y/N。AutoPilot は auto-Y。N → Cancelled ModuleResult が自動 return される | 必須 |
| Step 5 | **適用ループ** — try/catch で 1 件ずつ処理し `$successCount / $skipCount / $failCount` を集計 | 必須 |
| Step 5.5 | **Post-Apply Verification** — システム状態を読み返して期待値と一致するか検証。`$verified` を Step 6 に渡す。reg_hklm_config / firewall_config / hostname_config が参考実装 | 推奨 |
| Step 6 | **集計と return** — `New-BatchResult -Success N -Skip N -Fail N -Title "..."` で ModuleResult を構成して返す。`-Verified $verified` を付けると Status とは別に検証結果が記録される | 必須 |

冒頭には `Show-Separator` + `Write-Host -ForegroundColor Cyan` でモジュール名を出すヘッダブロック、オプションの P/Invoke `Add-Type` 用ブロック (使わなければ削除) も用意されている。

このフローは **fabriq モジュールの構造的同型性**を保証する。Claude が新規モジュールを書くときは Step 1〜6 のスケルトンに肉付けするだけで、`Initialize-Module` / `New-ModuleResult` / `New-BatchResult` / `Show-Info|Error|Success|Skip` / `Confirm-ModuleExecution` といった common.ps1 の API が自動的に正しい順番で呼ばれる。

### `_template_list.csv`

```
Enabled,TargetName,Description,Segment
1,example_item_1,First example item,
0,example_item_2,Second example item (disabled),
```

最小スキーマの例。実モジュールではここに必要な列を追加する。`Segment` 列は省略可で、Profile から segment 指定で呼び出されたとき一致行 + 空欄行のみが処理される。

### `module.csv`

```
MenuName,Category,Script,Order,Enabled
Template Module,System,_template_script.ps1,99,1
```

`fabriq_operator` の Modules タブ表示用メタデータ。MenuName / Category / Script / Order / Enabled の 5 列固定。

### `Guide.txt`

日本語で書かれたテンプレートの読み方。「コピー → リネーム → 編集」の手順、CSV 各列の意味、7 ステップの説明を記載。Claude / 開発者へのガイド。

### `VERSION` / `REQUIRES_KERNEL`

- `VERSION` = `0.1.0` (固定)。これは「**開発中・未リリース**」の目印で、`dev/template/` 配下にいる間は 0.1.x、`modules/` に配備された瞬間に新規モジュールとして 1.0.0 に書き換える運用 (CLAUDE.md ルール H)
- `REQUIRES_KERNEL` = `2.0.0` (現行 baseline)。新規モジュールが新規 API に依存する場合のみ昇格

---

## framework_overlay_rules.json

contracts に既述。`profiles/`, `kernel/csv/hostlist.csv`, `kernel/csv/categories.csv` 等を **overlay 時の保持対象**として宣言する単一ソース。`build_framework_patch.ps1` および外部の `fabriq_studio` の双方が同じファイルを参照することで、配布動作の一貫性を保証する。

---

## build_framework_patch.ps1

### Role
fabriq ソースツリーを output ディレクトリにミラーしつつ、site 固有 CSV / 実行時アーティファクト / `profiles/` ツリーを **除外** (上記 overlay rules に従う)。配布先で既存の現場データを上書きせずに重ねて適用できる「framework patch」を生成する。

### Trigger / 実行例
```powershell
powershell.exe -File .\dev\build_framework_patch.ps1
powershell.exe -File .\dev\build_framework_patch.ps1 -OutDir D:\share\patches
powershell.exe -File .\dev\build_framework_patch.ps1 -PatchName my-patch -Purpose "RC1"
```

### Output
- `<OutDir>/fabriq_patch_{yyyy-MM-dd}_kernel-v{KERNEL_VERSION}/` フォルダ
- 中に `PATCH_README.md` を自動生成 (含むファイル一覧、Purpose、kernel/モジュールバージョン)
- 配布先で「上から zip 解凍」で適用する想定

### 用法
**framework 全体配布版**。小さな差分パッチ (個別ファイルだけ) は手動で配布先のディレクトリ構造をミラーしてコピーする方式 (CLAUDE.md memory: `feedback_patch_creation.md` のとおり、targeted vs framework の二択で運用)。

---

## seed_module_versions.ps1

### Role
`modules/standard/` および `modules/extended/` 配下の全モジュールに対し、欠損している `VERSION` (=`1.0.0`) と `REQUIRES_KERNEL` (=`2.0.0`) を **idempotent に**作成する一括 seeder。既存ファイルは保持される。

### History
2026-04-23 に一斉実施済み。当初は lazy-seed (Claude が初めて touch した時に 1.0.0 打刻) だったが、fabriq_studio の update 機能で「両側 VERSION 欠損 = 比較不能 → SKIP だが実体差分あり」という現実的問題が起きたため、batch seed に切替。

### Trigger / 実行例
```powershell
powershell.exe -File .\dev\seed_module_versions.ps1            # 実書き込み
powershell.exe -File .\dev\seed_module_versions.ps1 -DryRun    # 確認のみ
```

### Output
- 各モジュール配下に欠損していた `VERSION` / `REQUIRES_KERNEL` を作成
- コンソールに「Seeded / Skipped / Already-present」のサマリを出力

---

## check_version.ps1

### Role
`kernel/KERNEL_VERSION` を真のソースとして、以下 3 箇所の版表記との整合を検証する:

- `README.md` L1 `# Fabriq ver{X.Y}` (major.minor)
- `kernel/common.ps1` L2 `# Easy Kitting Batch - Common Function Library v{X.Y.Z}` (full)
- `kernel/main.ps1` L3 `# Fabriq ver{X.Y}` (major.minor)

### Trigger / 実行例
```powershell
pwsh ./dev/check_version.ps1
```

### Output / 終了コード
- 0: 全一致
- 1: いずれかが不一致 (リリース手順の最後で必ず実行することが CLAUDE.md ルール K で義務付け)

---

## verify_comments_only.ps1

### Role
`.ps1` ファイルの変更が **コメント変更のみ**であることを PowerShell 自身のパーサで証明する検証ツール。トークン列を抽出し、Comment / NewLine 以外のトークン (Kind + Text) が完全一致すれば PASS。

### 使いどころ
- 日本語コメント → 英語コメントへの翻訳作業 (CLAUDE.md memory: `feedback_scripts_english_only.md` 由来) の安全性証明
- ドキュメントのみの修正であることを証明したいとき
- Comment-Based Help dot-keyword (`.SYNOPSIS` 等) は Comment トークン内なので検出できない (翻訳者が手で残すルール)

### Trigger / 実行例
```powershell
# 任意 2 ファイル比較
pwsh ./dev/verify_comments_only.ps1 -Original kernel/common.ps1.bak -Modified kernel/common.ps1

# git HEAD との比較
pwsh ./dev/verify_comments_only.ps1 -Path kernel/common.ps1
```

### 終了コード
- 0: PASS (コメントのみ差分)
- 1: FAIL (機能トークンに差分あり、または parse error)

NewLine トークンは **意図的に**比較対象外。LF/CRLF 差や `git show` の trailing newline 由来の偽 FAIL を避けるため。

---

## launcher/ — Fabriq.exe / Fabriq_IOS.exe

### Role
`Fabriq.bat` と同じく `kernel/main.ps1` を起動するだけの C# 極小ラッパー。**カスタムアイコン**と **UAC 自動昇格マニフェスト**を埋め込み、エクスプローラー / タスクマネージャーから「アプリ」として見せる。

### File 構成
| ファイル | 役割 |
|---|---|
| `Launcher.cs` | C# ソース。conhost + powershell.exe で `kernel\main.ps1` 起動 |
| `Launcher_IOS.cs` | fabriq_ios 版 (Fabriq_IOS.exe を生成) |
| `app.manifest` | UAC `requireAdministrator` 指定 (ダブルクリックで UAC ダイアログ) |
| `app_ios.manifest` | fabriq_ios 用マニフェスト |
| `fabriq.ico` | アイコン (初回ビルド時に shell32.dll から仮アイコン抽出) |
| `build.ps1` / `build_ios.ps1` | csc.exe を呼んで `..\..\Fabriq.exe` / `Fabriq_IOS.exe` を生成 |

### ビルド
```powershell
cd e:\fabriq\dev\launcher
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

### 依存
.NET Framework 4.x 付属の `csc.exe` のみ (Windows 10/11 標準で揃う)。管理者権限不要。

### 再ビルド要否
- Launcher.cs / app.manifest / fabriq.ico を変更した場合のみ
- fabriq 本体 (kernel / modules / main.ps1) を変更しても **再ビルド不要** (ラッパーは変えない)

---

## ico/ — アイコン素材

`app_icon.ico`, `app_icon_preview.png`, JPG 素材 2 枚。launcher 用ではなく、fabriq_evidence_manager / fabriq_studio など他プロジェクトでの共通使用想定。


<!-- ============================================================ -->
# === apps/05_commands.md ===
<!-- ============================================================ -->

# commands/ — 手動操作コマンド集

`e:/fabriq/commands/` は、**Status Monitor の手動操作**または **System Launcher 経由**で操作員が任意のタイミングで呼び出すユーティリティスクリプト群です。Profile 経由で自動実行されるモジュールとは異なり、ここに置かれるスクリプトは「操作員が必要だと判断したとき」に走らせる手元の道具箱としての位置付けで、エビデンス採取や履歴記録の対象外であることが多いのが特徴です。

## ファイル一覧

| ファイル | 役割 | 起動経路 |
|---|---|---|
| `gpupdate_command.ps1` | グループポリシー強制更新 | Status Monitor / System Launcher |
| `temp_command.ps1` | カスタム差し込み用テンプレート (空) | Status Monitor |
| `explore_restart_command.ps1` | Explorer 再起動 | Status Monitor |
| `diag_crypto.ps1` | 暗号化 / passphrase 状態の診断 | Status Monitor / 開発者手動 |
| `get_evidence.ps1` | PC 情報の手動採取 (split log 形式) | Status Monitor / FabriqApps |
| `system_launcher.ps1` | Windows 設定ショートカットパレット (apps/system_launcher と同一実装) | Status Monitor |

---

## gpupdate_command.ps1

### Role
`gpupdate /force` を実行してグループポリシーを即時適用する。ドメイン参加直後に GPO ベースのポリシーを反映させたい場合に手動で呼ぶ。

### 中身
- `gpupdate /force` を try/catch で実行
- 結果に応じて [SUCCESS] / [ERROR] を色付き出力
- 末尾 `pause` でユーザーが結果を確認してから閉じる

### 起動コンテキスト
Status Monitor の手動ボタンから、または System Launcher の管理ツール枠経由。Profile に組み込むケースは少ない (組み込むなら専用モジュール化される)。

---

## temp_command.ps1

### Role
**空のテンプレート**。中身は `# ========================================` の枠 + `# template` の 1 行のみ。現場ごとに必要な手作業をその場で書き込んで使う「使い捨てスロット」。

### 中身
```
# ========================================
# template
# ========================================
```

### 起動コンテキスト
Status Monitor の手動枠に「temp」ボタンとして固定で表示される。ここに案件固有の臨時処理を書き込んで実行する想定。

---

## explore_restart_command.ps1

### Role
`explorer.exe` を停止 → Windows の自動再起動を 15 秒まで待機 → 起動したか確認するコマンド。レジストリ変更の即時反映、Start レイアウト適用後のタスクバー / デスクトップ再描画などに使う。

### 中身
- `Confirm-Execution` で Y/N 確認 (common.ps1 経由)
- `Stop-Process -Name explorer -Force`
- 1 秒間隔で 15 秒まで起動を polling
- 復帰しなかった場合は警告

### 起動コンテキスト
Status Monitor。デスクトップやタスクバーが一時的に消える操作なので、ユーザー確認必須。

---

## diag_crypto.ps1

### Role
fabriq の暗号化機能 (`Unprotect-FabriqValue` / DPAPI / passphrase) が正しく動いているかを **診断**する。CSV 内に `ENC:` プレフィックスの暗号化値があるかを scan し、復号成否を試す。

### 中身
1. `$global:FabriqMasterPassphrase` の有無確認
2. `Unprotect-FabriqValue`, `Import-ModuleCsv`, `Import-CsvSafe` 関数の存在確認
3. fabriq 配下の `*.csv` を再帰スキャンして `ENC:` パターンを含むファイルを列挙
4. 各 ENC 値に対して復号試行、成否を表示

### 起動コンテキスト
- 開発者の手動診断
- Status Monitor からトラブルシュート用に呼ぶ (パスフレーズ未投入で `[ENCRYPTED]` のまま動いてしまう問題の切り分け)

CLAUDE.md memory: `project_crypto_security_review.md` で挙げた懸念の確認用。

---

## get_evidence.ps1

### Role
PC 情報を手動採取して `evidence/pc_information/` 配下にテキストログとして保存する。`evidence_config` モジュールが定型化したエビデンス採取とは別経路の **緊急時 / 個別採取**用。

### 中身
- 出力先: `$global:FabriqEvidenceBasePath\pc_information` (未設定時は `evidence/pc_information/<COMPUTERNAME>_<yyyyMMdd_HHmmss>/`)
- マスターログ `_ALL_<COMPUTERNAME>_Log.txt` + セクション別 split log (例: `01_SystemInfo.txt`)
- `Out-Log` 関数で「Console / Master / Split」3 箇所同時出力
- `Start-Section -Title -FileName` でセクション切替
- 採取対象: 基本情報 (hostname / OS / spec)、その他 PC 構成

### 起動コンテキスト
Status Monitor、または FabriqApps から「evidence_config が動かなかったとき / その前後で個別採取したいとき」用に手動起動。

---

## system_launcher.ps1

### Role
`apps/system_launcher/system_launcher.ps1` と **同一実装**。commands/ にも置かれているのは、Status Monitor から手動操作枠で 1-click 起動できるようにするため。

### 中身
- 34 項目の Windows ツール (ms-settings: 系 / *.cpl / *.msc / shell:::{GUID} / cmd / powershell / runas)
- `Invoke-Tool` で Type に応じた Start-Process 起動分岐

### 起動コンテキスト
Status Monitor の手動ボタン枠。apps/ 側の利用と機能は同じ。


<!-- ============================================================ -->
# === profiles/00_profiles_overview.md ===
<!-- ============================================================ -->

# profiles カタログ

`e:/fabriq/profiles/` には fabriq の Profile (= モジュール実行手順書 CSV) のサンプル / テンプレートが収められています。実運用では各 Profile は **現場固有の編集物**であり、`framework_overlay_rules.json` により `profiles/` ツリー全体が overlay 時に **保持対象 (preserved)** に指定されています。つまりここに置かれているファイルは「あくまで出発点」で、実際のキッティング案件では現場で書き換えられて使われます。

## Profile CSV のスキーマ

すべての Profile CSV は最低限 `Order,ScriptPath,Enabled,Description,Segment` の 5 列を持ちます (`Segment` は省略可)。一部の Profile (sysprep.csv, _test_harness.csv) は `Note` や `ErrorMode` 列を追加で持ち、3.2.0 以降の FlexProfile 対応 Profile は `Group` 列を持つこともあります。

特殊マーカー:

- `__RESTART__` — 再起動を挿入。fabriq は再起動後 RunOnce で resume する
- `__AUTOPILOT__` — AutoPilot モード ON、`Description` に `WaitSec=N` を書くと wait 秒数指定可
- `__ASYNC__` — このマーカー以降のモジュールを非同期 (Runspace) 実行に切替

## 同梱されている Profile

### Master_Pre01.csv (1 行 / 単発)

```
10,standard/driver_config/driver_export_config.ps1,1,Driver Export
```

**シナリオ**: マスター機 (golden master) からドライバを **エクスポート**する単発作業。Pre01 は「マスター機作成側」で 1 度だけ走らせる位置付け。

### Master_Pre02.csv (9 行)

```
10  builtin_admin_config             ローカル Administrator 有効化など
20  autologon_config                 自動ログオン設定
30  __RESTART__
40  local_user_delete                既存ローカルユーザー整理
50  reg_hklm_config / 60 reg_hkcu_config
70  scheduled_task_disable_config
80  bitlocker_disable
90  driver_import_config             Pre01 でエクスポートしたドライバを適用
```

**シナリオ**: 配布先 PC の **キッティング前段階** (OS 入れ替え直後)。再起動を挟んで管理者・自動ログオン・レジストリベースライン・ドライバ適用までを行う。

### Master_Config01.csv (17 行)

ライセンスアクティベーション → 壁紙 → 時刻同期 → ファイアウォール → 電源 → ローカルユーザー作成 → レジストリ HKLM/HKCU → アプリインストール (winget) → ODT → DPI → 解像度 → ファイルコピー → プリンタドライバ / プリンタ登録、と続く **本体構成の主要 Profile**。

**シナリオ**: キッティングの中核作業。1 台あたり最も時間がかかるフェーズ。

### Master_Config02.csv (5 行)

```
10  startlayout_backup_config
20  startlayout_build_config
30  default_app_config / export_app_associations
40  desktop_icon_backup
50  fabriq_app_launcher
```

**シナリオ**: マスター機側で **スタートレイアウト・既定アプリ・デスクトップアイコン**を採取して `app_associations.xml` 等を作る。Pre01 と同じく golden master 系の作業。

### Master_Config03.csv (5 行)

```
10  taskbar_config
20  startlayout_import_config
30  startlayout_delete_config
40  desktop_icon_restore
50  default_app_config
```

**シナリオ**: Config02 で採取した成果物を配布先で **インポート / 復元**する Profile。

### Master_Config04.csv (5 行)

```
10  evidence_config           証跡採取
20  directory_cleaner         不要ディレクトリ消去
30  history_destroyer         履歴/キャッシュ消去
40  storeapp_config           Store アプリ削除
50  sysprep_config            Sysprep 構成
```

**シナリオ**: キッティング **完了直前 / 出荷前**の仕上げ。証跡を採ってから掃除して Sysprep。

### Custom Plan.csv (5 行)

```
10  reg_hklm_config / 20 reg_hklm_delete
30  __RESTART__
40  reg_hkcu_config / 50 reg_hkcu_delete
```

**シナリオ**: 案件固有のレジストリ調整を入れる **カスタムスポット**枠。再起動を挟んで HKLM 系と HKCU 系を分離。

### sysprep.csv (9 行)

`Order,ScriptPath,Enabled,Description,Segment,Note,ErrorMode` の拡張スキーマ。app_config (Enabled=0)、reg_hklm_delete、storeapp_config、taskbar_config、default_app_config、generic_process_runner、`__RESTART__` (Enabled=0)、history_destroyer、sysprep_config の順。

**シナリオ**: 出荷直前の **Sysprep 専用ライン**。`__RESTART__` をオフにしておくのは、sysprep_config 自体が再起動 / シャットダウンを最終ステップで担うため。

### _test_harness.csv (11 行)

`__AUTOPILOT__` で WaitSec=1 にした上で、`test_harness_config` モジュールの各セグメント (success_verified, success_verifail, success_no_verify, skipped, partial, cancelled, error_basic, retry_success, retry_exhaust) を順に呼ぶ統合テスト Profile。`__RESTART__` を中盤に挟んで resume 動作の検証もする。

**シナリオ**: kernel / FlexProfile / Status Monitor のリグレッションテスト。各 Status badge 描画と ErrorMode (skip / retry) の挙動、resume 動作を 1 本で検証する。

### _test_harness_async.csv (7 行)

`__AUTOPILOT__` + `__ASYNC__` を併用し、非同期実行パスのテストを行う。`hang_sim` セグメントで Status Monitor の手動 Skip も検証対象。

**シナリオ**: 非同期 Runspace 経路のリグレッションテスト。

### easy_template/ (3 ファイル)

```
easyprofile.bat       管理者昇格 + powershell -File easyprofile.ps1
easyprofile.ps1       AutoPilot 軽量ランナー (history なし / evidence なし / checklist なし)
easyprofile.csv       Enabled,Script,Description の 3 列 (Order なし、上から順実行)
```

**シナリオ**: fabriq の正規ダッシュボードを通さずに「**選択された数モジュールだけ即実行する 1-shot ランナー**」を作るためのテンプレート。`profiles/easy_profile_<X>/` という規約で複数並べて運用できる。Hostname 設定 / ローカルユーザー作成 / 削除のような単発タスクを、エビデンス無しで素早く流すユースケース。デフォルト 3 行はすべて Enabled=0 で安全側。

## 運用ルール

- `framework_overlay_rules.json` により `profiles/` ツリー全体が overlay 時に保持される (= フレームワークアップデートで現場 Profile が消えない)
- 同梱されているこれらのファイルは「**出発点 (starting points)**」であり、現場では命名 / Order / Enabled / Segment を改変して使われる
- `_test_harness*.csv` は test_harness_config モジュールに依存するため、本番現場の Profile としては使わない (開発用)
- `easy_template/` はディレクトリごとコピーして `profiles/easy_<案件名>/` の形で運用するためのスケルトン


<!-- ============================================================ -->
# === modules/00_modules_overview.md ===
<!-- ============================================================ -->

# モジュール全体図 — 標準 60 / 拡張 15

fabriq の機能はすべて `modules/{standard,extended}/<name>/` に packaging されている。各モジュールは独立 SemVer で配備され、要求カーネル版を `REQUIRES_KERNEL` で宣言する。

カテゴリは `kernel/csv/categories.csv` で定義され、ダッシュボードのグルーピング順序を決定する。

詳細な per-module 解説は本ディレクトリの `<name>.md` 各ファイルを参照。

---

## Standard（60 件）

### Network

| モジュール | 主な役割 | Verification |
|---|---|---|
| `hostname_config` | PC 名変更（再起動で反映） | あり（pending value 検証） |
| `ipaddress_config` | 静的 IP / Subnet / Gateway / DNS 設定 | あり（読み返し） |
| `temp_ipaddress_config` | プールから一時 IP 取得（GUI 選択 + DAD） | あり |
| `domain_join` | ドメイン参加（`ENC:` パスワード） | 不可（再起動後反映） |
| `ssid_config` | Wi-Fi プロファイル追加（netsh wlan add profile） | あり |

### Display

| モジュール | 主な役割 | Verification |
|---|---|---|
| `brightness_config` | 輝度設定（WMI） | あり |
| `dpi_api_config` | DPI スケール（Win32 API） | 不可 |
| `resolution_api_config` | 解像度変更（ChangeDisplaySettings） | 部分 |

### Desktop

| モジュール | 主な役割 | Verification |
|---|---|---|
| `wallpaper_config` | 壁紙（SystemParametersInfo） | なし |
| `taskbar_config` | LayoutModification.xml を Default User へ | なし |
| `startlayout_config` | スタートメニュー Backup/Build/Import/Delete pair | 部分 |

### Security

| モジュール | 主な役割 | Verification |
|---|---|---|
| `bitlocker_config` | BitLocker 有効化（async, await pair） | 不可（async 完了は別） |
| `firewall_config` | プロファイル別 ON/OFF + ルール設定 | あり |
| `firewall_rule_config` | ルール export / import pair | あり（両方） |
| `firewall_rule_make_config` | ルール手動定義 | あり |
| `cert_config` | 証明書ストア配置（`ENC:` パスワード） | あり |
| `office_license_config` | Office キー install + Activate pair | de-facto（auth が verify） |
| `windows_license_config` | Windows キー install + Activate pair | 部分（Activate のみ） |

### User Management

| モジュール | 主な役割 | Verification |
|---|---|---|
| `local_user_config` | ローカルユーザ作成 + 削除 pair | あり（作成のみ） |
| `profile_delete` | ユーザプロファイル削除 | なし（推奨実装あり、未実装） |

### Printer

| モジュール | 主な役割 | Verification |
|---|---|---|
| `printer_driver_config` | ドライバ install + register + uninstall trio | あり（install + register） |
| `printer_delete` | プリンタ削除（printer_list.csv 参照） | あり |

### Applications

| モジュール | 主な役割 | Verification |
|---|---|---|
| `app_config` | EXE/MSI/MSU インストール（共有 CSV） | 不可（複雑） |
| `winget_install` | winget update / install / upgrade trio | なし（winget 自身で完結） |
| `bloatware_remove` | UWP / Provisioned / Capability 削除 | あり |
| `bloatware_export` | 現状の bloatware 一覧 export | N/A（read-only） |
| `storeapp_config` | StoreApp 一括削除 | あり |
| `odt_config` | ODT で Office インストール | なし |
| `browser_addon_config` | Edge / Chrome 拡張機能 ADMX 経由 | あり |
| `fabriq_app_launcher` | apps/ ランチャ（FabriqApps ボタン本体） | N/A |

### Power

| モジュール | 主な役割 | Verification |
|---|---|---|
| `power_config` | 電源プラン（P/Invoke powrprof.dll、HP OEM 対策で Win32 API） | あり |

### Maintenance

| モジュール | 主な役割 | Verification |
|---|---|---|
| `acl_config` | ACL backup + restore pair | 除外（誤 PASS リスク） |
| `copyfile_config` | ファイル / ディレクトリコピー | 除外（誤 PASS リスク） |
| `file_delete` | ファイル / ディレクトリ削除 | あり |
| `office_update` | Office 更新（OffScrubC2R or click-to-run update） | de-facto |
| `partition_config` | パーティション作成 / 拡張 | あり（±5%） |
| `robocopy_config` | UNC / Local の robocopy（`ENC:` パスワード） | なし |
| `system_finalize` | shell32 reload + cache flush + Explorer 再起動 | なし |

### System

| モジュール | 主な役割 | Verification |
|---|---|---|
| `autologon_config` | AutoLogon registry 設定（`ENC:` パスワード）| 不可（再起動で反映） |
| `default_app_config` | DISM / SetUserFTA で既定アプリ設定 | 不可（新プロファイルで反映） |
| `driver_config` | DISM driver export + import pair | 不可 |
| `generic_process_runner` | EXE / MSI を generic に発火 | なし |
| `ppkg_config` | プロビジョニングパッケージ install + uninstall | なし |
| `process_killer` | プロセス強制終了（idempotent） | 意図的になし |
| `restart_config` | 再起動（AutoPilot 非対応） | 不可 |
| `restore_point` | システム復元ポイント作成 | 部分 |
| `scheduled_task_config` | タスクスケジューラ enable/disable pair（共有 CSV）| 部分 |
| `signout_config` | サインアウト | 不可 |
| `spi_config` | SystemParametersInfo + Active Setup（HKCU 系） | 除外 |
| `sysprep_config` | unattend.xml + SetupComplete.cmd 生成 | 除外 |
| `time_sync_config` | w32tm NTP 設定 + sync source verify | あり（retry 込） |
| `volume_config` | Core Audio API でマスターボリューム | あり |

### Registry

| モジュール | 主な役割 | Verification |
|---|---|---|
| `reg_hklm_config` | HKLM レジストリ + delete pair（Test-RegistryValueMatch 共有関数） | あり |
| `reg_hkcu_config` | HKCU + Default Profile + Active Setup/Startup Batch + delete pair | あり |

### Scripts

| モジュール | 主な役割 | Verification |
|---|---|---|
| `generic_batch_runner` | 任意 .bat / .ps1 / .cmd を generic 実行 | なし |
| `startup_command_config` | Default User Startup folder にコマンド配置 | なし |

### Evidence

| モジュール | 主な役割 | Verification |
|---|---|---|
| `evidence_config` | 22 セクションのシステム情報収集 + manifest.json 生成 | N/A |

### Test

| モジュール | 主な役割 | Verification |
|---|---|---|
| `test_error_module` | ErrorMode 検証用エラー発生器 | N/A |
| `test_harness_config` | マルチシナリオテストハーネス | あり（by design） |

### Standalone（Profile 直接登録不可）

| モジュール | 主な役割 |
|---|---|
| `windows_update` | スタンドアロン COM-API ループ。`module.csv` 無し、`Invoke-WindowsUpdateLoop` から `[wu]` 経由起動 |

---

## Extended（15 件）

### Network

| モジュール | 主な役割 | Verification |
|---|---|---|
| `ipv6_config` | IPv6 binding toggle（adapter ごと） | なし |
| `network_profile_config` | NetworkList Category 書き換え（baseline + override パターン） | あり |

### Display

| モジュール | 主な役割 | Verification |
|---|---|---|
| `display_config` | PrimSurfSize.cx/cy registry（マルチモニタ）| なし（再起動要） |
| `dpi_config` | per-monitor DPI（embedded C# DpiScaleResolver、HKCU + Default-hive dual-write）| なし |

### Desktop

| モジュール | 主な役割 | Verification |
|---|---|---|
| `desktop_icon_config` | Desktop アイコン .reg backup + restore pair（HKCU↔HKU\<SID> 正規化）| なし |

### User Management

| モジュール | 主な役割 | Verification |
|---|---|---|
| `builtin_admin_config` | Built-in Administrator アカウント設定（Enabled 列で Enable/Disable）| なし |
| `group_config` | ローカルグループメンバ管理（WMI で CurrentUser 解決）| なし |

### Maintenance

| モジュール | 主な役割 | Verification |
|---|---|---|
| `directory_cleaner` | ディレクトリクリーンアップ（hardcoded forbidden-path whitelist 3 重 guard）| なし |
| `history_destroyer` | 13 カテゴリ履歴削除（Edge / Chrome / Search / Wi-Fi / 7 special handlers）| なし |

### System

| モジュール | 主な役割 | Verification |
|---|---|---|
| `azure_ad_join_check` | dsregcmd 出力解析（read-only、script_looper retry 設計）| N/A |
| `reg_template` | レジストリ template backup / import pair（汎用テンプレ）| なし |

### Scripts

| モジュール | 主な役割 | Verification |
|---|---|---|
| `script_looper` | OnError / Always retry framework（pipeline + global 二重 ModuleResult 検出）| なし |

### ManualWorks

| モジュール | 主な役割 | Verification |
|---|---|---|
| `manual_kitting_assistant` | 手動作業アシスタント WinForms（Gundam Light theme、NoActivate window）| N/A（Enabled=0 デフォ） |

### Evidence

| モジュール | 主な役割 | Verification |
|---|---|---|
| `log_uploader` | robocopy で UNC / Local へエビデンス + ログ転送（log_destinations.csv 参照）| 不可（best-effort） |

### Special: Pianist (extended)

| モジュール | 主な役割 | Verification |
|---|---|---|
| `pianist` | **マルチフェーズ GUI maestro**。3-tab Phase view（Procedure / Samples / Values）+ section markers + modeless image viewer + Pause/Stop/Speed (v1.6.0+)。apps/ から extended/ へ昇格（2026-05-02）、2189 行の独立色濃いモジュール | あり（Manual phase status aggregation） |

---

## モジュール構成ファイル（共通）

各モジュールは以下のファイル群で構成される：

| ファイル | 役割 | overlay 区分 |
|---|---|---|
| `module.csv` | メニュー名・カテゴリ・表示順・有効無効（複数行可） | framework |
| `<name>.ps1` | 実行スクリプト本体（dev/template ベース） | framework |
| `<other>.ps1` | 補助スクリプト（_install / _uninstall / _backup / _restore 等） | framework |
| `<name>_list.csv` | 設定データ（対象リスト等） | site-specific |
| `Guide.txt` | 使い方ガイド（日本語） | framework |
| `preset.csv` | Studio 用ドロップダウン UI 定義（任意） | framework |
| `VERSION` | モジュール SemVer（1 行 X.Y.Z） | framework |
| `REQUIRES_KERNEL` | 要求最小カーネル版（1 行 X.Y.Z） | framework |

---

## カテゴリ統計

| カテゴリ | Standard | Extended |
|---|---|---|
| Network | 5 | 2 |
| Display | 3 | 2 |
| Desktop | 3 | 1 |
| Security | 7 | 0 |
| User Management | 2 | 2 |
| Printer | 2 | 0 |
| Applications | 8 | 0 |
| Power | 1 | 0 |
| Maintenance | 7 | 2 |
| System | 14 | 2 |
| Registry | 2 | 0 |
| Scripts | 2 | 1 |
| Evidence | 1 | 1 |
| Test | 2 | 0 |
| ManualWorks | 0 | 1 |
| Standalone | 1 | 0 |

---

## Verification 実装率

- **実装済み**: 25 / 75 ≈ 33%
- **意図的除外**（誤 PASS リスク or 検証不可能）: 約 15 件
- **未実装（推奨）**: 残り

検証除外リスト（feedback memory `project_verification_exclusions`）:
- `acl_config`: ACL ツリー完全読み返しは膨大、サブセット検証で false PASS の risk
- `spi_config`: Default Profile への hive load 経由でログイン後にしか反映されない
- `copyfile_config`: ファイル存在 != 内容正しい、ハッシュ検証は重い
- `sysprep_config` / `restart_config` / `signout_config` / `domain_join`: 再起動後 / OS 再起動後 / 別ユーザログオン後にしか確認できない（技術的不可）


<!-- ============================================================ -->
# === modules/acl_config.md ===
<!-- ============================================================ -->

# acl_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: ACL Backup / ACL Restore
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 不可（誤 PASS リスクのため意図的に非実装）
**サブスクリプト**:
- `acl_backup.ps1` … `icacls /save /T` によるディレクトリ ACL のフルバックアップ + 継承断ち subdir の個別バックアップ
- `acl_restore.ps1` … フル復元 → 浅い順に個別 override の二段復元

## 目的
キッティング作業や移行作業に伴ってファイル ACL を巻き戻す必要が出た場合に備え、対象ディレクトリツリーの ACL をハイブリッド方式（フル `/save /T` ＋ 継承断ち subdir の個別保存）でバックアップ／レストアします。フルセーブのみでは継承断ちフォルダの ACL がわずかにずれ、サブディレクトリを 1 つずつ保存すると大規模ツリーで激遅になるという `icacls` の二律背反を、ハイブリッド方式で解消します。冪等性の観点では、復元時はフル → 浅い順 → 深い順という順序を守ることでより具体的な ACL が必ず後勝ちで適用される設計になっています。

## 入力 (CSV)
`acl_list.csv` の主な列:
- `Enabled` … 1=処理 / 0=スキップ
- `Id` … バックアップフォルダ名（`backup/{Id}_{SafeName}/`）に使う一意識別子
- `TargetPath` … バックアップ／復元の起点ディレクトリ。`%USERPROFILE%` `%SELECTED_NEW_PCNAME%` 等の環境変数を `Expand-UserEnvironmentVariables` で展開
- `Description` … 表示ラベル
- `Segment` … `Import-ModuleCsv` による Segment フィルタ用（空欄は常に採用）

## 主要ステップ
1. `acl_list.csv` 読み込み（`Import-ModuleCsv -FilterEnabled` で Segment フィルタ込み）
2. Pre-flight: `Test-AdminPrivilege` と `icacls.exe` 存在確認
3. Pre-execution display（dry-run 一覧）
4. 実行確認（AutoPilot 時は自動 Y）
5. Backup: フル `/save /T` を撮ったあと、継承断ち subdir を SHA256 ハッシュ名で個別保存し `_manifest.csv` を出力 / Restore: フル復元 → manifest を浅い順にソートして個別 override
6. 結果集計（`New-ModuleResult`）

## 注意点・運用メモ
- **管理者権限必須**（`icacls /save` `/restore` ともに admin token が必要）
- バックアップ出力構造は `backup/{Id}_{SafeName}/_full_acl.txt`, `_manifest.csv`, `individual/{hash}_acl.txt`
- シンボリックリンク／ジャンクションはスキャン対象外
- Restore は事前 Backup が無いと Error
- 継承断ち subdir が無い場合はフル復元のみで完了
- `_manifest.csv` がパス→ハッシュ名のマッピング表として残るので、ハッシュファイルから元パスへ逆引き可能

## 検証
本モジュールは Post-Apply Verification を **意図的に非実装** としています。これは fabriq プロジェクトルール（CLAUDE.md「Post-Apply Verification 除外モジュール」に準拠）の判断で、`icacls /save` のテキスト出力と現在の ACL を厳密に比較するには継承フラグ／SID 解決／ACE 順序／差分検出を完全に再現する必要があり、誤 PASS（false positive）リスクが過大なためです。「PASS と表示されたが実際は不一致」という最悪ケースを誘発しないよう、`-Verified` を渡さず Verified 列は空欄になります。検証は `icacls /save` の出力テキストを目視差分するか外部ツールで実施することが推奨されます。


<!-- ============================================================ -->
# === modules/app_config.md ===
<!-- ============================================================ -->

# app_config (Standard)

**カテゴリ**: Applications
**メニュー名**: App Installation
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（インストーラー ExitCode のみで成否判定）
**サブスクリプト**: なし（メイン `app_config.ps1` のみ。インストーラー本体は `file/` 配下に配置）

## 目的
キッティング工程で導入する任意のサードパーティアプリケーション（Chrome / Acrobat Reader / 7-Zip など）を、CSV に記述された順序で一括サイレントインストールするためのモジュールです。インストーラー実体は `file/` ディレクトリに配置し、CSV の `FileName` 列がそれを参照します。`exe` / `msi` / `bat` の 3 種別を統一インターフェースで扱えるため、ベンダーごとに違うサイレントオプションをすべて CSV に集約できます。

## 入力 (CSV)
`app_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `AppName` … アプリ識別名
- `FileName` … `file/` 内のインストーラーファイル名
- `Type` … `exe` / `msi` / `bat`
- `SilentArgs` … サイレントインストール用引数文字列
- `Description` … 表示用ラベル（指定時は AppName の代わりに表示）
- `Segment` … Segment フィルタ

`Type` ごとの実行形態:
- `exe` … `& <Installer> <SilentArgs>`
- `msi` … `msiexec /i <Installer> <SilentArgs>`
- `bat` … `cmd /c <Installer>`

## 主要ステップ
1. `app_list.csv` を読み込み、Enabled=1 の行を抽出
2. `file/` ディレクトリと CSV の `FileName` 実在を突合
3. インストール対象の一覧表示（dry-run）
4. 実行確認（AutoPilot 時は自動 Y）
5. 各インストーラーを順次起動し ExitCode を判定（0=成功 / 3010=成功・再起動保留 / その他=失敗）
6. 結果集計を `New-BatchResult` で返却

## 注意点・運用メモ
- 多くのインストーラーは管理者権限を要求するため、Fabriq セッションを管理者で起動しておくこと
- **既インストール判定は実装していません**（毎回起動）。冪等性は各インストーラーのサイレントオプションに委ねる方針
- ExitCode 3010 は「成功・再起動保留」として Success 扱い。後続の `restart_config` で集約再起動する運用が想定
- MSI 実行には `msiexec.exe`、BAT 実行には `cmd.exe` が必要（標準 Windows 環境では常時 OK）
- `file/` 配下の実体配布は別途行う（GitHub にバイナリを置かない設計）

## 検証
Post-Apply Verification は **未実装**。アプリのインストール先パスや製品レジストリ存在は本モジュールでは検証しません。アプリ単位で検証要件が大きく異なる（HKLM レジストリにエントリを残すか／App-V／MSIX か等）ため、汎用検証ロジックを書くと誤判定が増える性質を持つためです。インストール痕跡の確認は `bloatware_export` で生成されるアプリインベントリ CSV を Evidence として残すことで間接的にカバーする運用です。


<!-- ============================================================ -->
# === modules/autologon_config.md ===
<!-- ============================================================ -->

# autologon_config (Standard)

**カテゴリ**: System
**メニュー名**: AutoLogon Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（実装すると平文パスワード露出 / AutoLogonCount 自動デクリメントのため意図的に非実装）
**サブスクリプト**: なし

## 目的
キッティング工程で「再起動を挟んでも一度だけ自動ログオンしたい」需要に応えるため、`HKLM\...\Winlogon` に `AutoAdminLogon` / `DefaultUserName` / `DefaultPassword` / `DefaultDomainName` / `AutoLogonCount=1` を書き込みます。`AutoLogonCount=1` を使うことで「次回 1 回限り」の挙動になり、ログオン後に Windows 自身が資格情報を自動消去するため、放置時のセキュリティリスクを最小化しています。Profile 実行時は `__AUTO_to_<User>__` マーカー経由で `$env:FABRIQ_AUTOLOGON_USER` が渡され、対象ユーザーが自動選択されます。

## 入力 (CSV)
`autologon_list.csv` の主な列:
- `Enabled` … 1=対象 / 0=スキップ
- `No` … 識別番号（対話 No 入力 / Profile 指定でも使用）
- `User` … ユーザー名
- `Password` … パスワード（**`ENC:` プレフィックス暗号化対応**。マスターパスフレーズで復号）
- `Domain` … ドメイン名（ローカルユーザーは空欄）
- `Description` … 表示ラベル
- `Segment` … Segment フィルタ

## 主要ステップ
1. `autologon_list.csv` を読み込み（`Import-ModuleCsv -FilterEnabled`）
2. `$env:FABRIQ_AUTOLOGON_USER` または対話 `No` 番号で対象ユーザー決定（Enabled=1 が 1 件なら自動選択）
3. 対象設定を表示（パスワードはマスク表示）
4. **冪等性チェック**: 現在のレジストリ値（`AutoAdminLogon=1` / `DefaultUserName` 一致 / `AutoLogonCount>=1`）を読み返し、一致なら Skipped で終了
5. 実行確認（AutoPilot 時は自動 Y）
6. レジストリ書き込み（5 値を `Set-ItemProperty -ErrorAction Stop` で一括設定）

## 注意点・運用メモ
- **管理者権限必須**（HKLM 書き込み。明示チェックはせず HKLM 書き込み失敗で検出）
- パスワードは `ENC:` プレフィックス対応。ENC 復号失敗時は対話モードで入力を求める／Error 終了
- `AutoLogonCount=1` を必ず付けるため、無人放置で何度も自動ログオンされる事故を防止
- ローカルアカウント運用時は `Domain` を空欄、ドメイン参加端末でドメインユーザー指定時は `Domain` を埋めること
- Profile 連携時は profile 側に `__AUTO_to_<User>__` を書き、kernel ランナーが `$env:FABRIQ_AUTOLOGON_USER` を立ててからこのモジュールを起動する流れ

## 検証
Post-Apply Verification は **意図的に非実装**:
1. `DefaultPassword` の比較は平文パスワードをログ／メモリに露出しうるため不可
2. `AutoLogonCount` は次回ログオン時に Windows 側で自動デクリメントされる一時値で、適用直後は 1、ログオン後は 0 のどちらも正常状態となり判定が曖昧
3. その代替として、Step 6 で `Set-ItemProperty -ErrorAction Stop` を使い `$failCount` をカウントすることで「書き込みが受理されたか」の実質検証としています

実際の自動ログオン挙動は次回再起動時の動作で確認する運用です（fabriq 標準では `restart_config` 経由）。


<!-- ============================================================ -->
# === modules/azure_ad_join_check.md ===
<!-- ============================================================ -->

# azure_ad_join_check (Extended)

**カテゴリ**: System
**メニュー名**: Azure AD Join Check
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 不可（状態確認のみで適用ステップが存在しない。Status 自体が検証結果を兼ねる）
**サブスクリプト**: なし（`azure_ad_join_check.ps1` 単体）

## 目的
`dsregcmd /status` を実行して Azure AD (Entra ID) Join が完了しているかを確認するモジュール。
未完了であれば Error を返すため、`script_looper` と OnError 条件で組み合わせることで「Join 完了まで自動で待ち続ける」ループを構築できる。
状態確認のみでシステムへの変更は一切行わない非破壊モジュールであり、`Confirm-ModuleExecution` も呼び出さず AutoPilot/手動どちらでも同一動作。

## 入力 (CSV)
なし。`dsregcmd /status` の出力から `AzureAdJoined` 値を自動判定する。

## 主要ステップ
1. `dsregcmd.exe` の存在確認（`%SystemRoot%\System32\dsregcmd.exe`）
2. `dsregcmd /status` を実行し標準出力を取得
3. 出力内 `AzureAdJoined : YES` を正規表現で検出して画面表示
4. YES なら Success、NO や検出不可なら Error を `New-ModuleResult` で返却

## 注意点・運用メモ
- 単体実行も可能だが、本来の用途は `script_looper` の `Condition=OnError` でリトライ駆動させること
- 管理者権限不要（`dsregcmd` は一般ユーザーで動作）
- Windows 10 以降前提（`dsregcmd.exe` が必要）

## 検証
適用ステップが無いため Post-Apply Verification は実装されない。`-Verified` を渡さないため履歴の Verified 列は空欄。Status 値（Success/Error）が「現在 Join 済みかどうか」をそのまま表す。


<!-- ============================================================ -->
# === modules/bitlocker_config.md ===
<!-- ============================================================ -->

# bitlocker_config (Standard)

**カテゴリ**: Security
**メニュー名**: BitLocker Enable / BitLocker Disable / BitLocker Await
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（暗号化／復号化は非同期で完了が後続するため、`bitlocker_await` で完了確認する設計）
**サブスクリプト**:
- `bitlocker_config.ps1` … BitLocker 有効化 + 回復キーをエビデンス保存（TPM-only / TpmAndPin 両対応）
- `bitlocker_disable.ps1` … BitLocker 無効化（復号開始）。FullyDecrypted は冪等スキップ
- `bitlocker_await.ps1` … Encrypting / Decrypting 中ドライブの完了をポーリング待機（30 秒間隔、30 分進行なしでタイムアウト）

## 目的
TPM 搭載 PC のキッティングで多用する「BitLocker 有効化 → 完了待機 → 必要なら無効化」を 3 スクリプトに分離して提供するモジュールです。回復キーは `evidence/bitlocker/<日付>_<UID>_<PC名>/<PC名>_<ドライブ名>.txt` に自動保存され、納品時のエビデンスとしてそのまま使えます。TPM-only と TPM+PIN（TpmAndPin Protector）の両対応で、PIN は ENC: プレフィックス暗号化または `hostlist.csv` 経由 (`$env:SELECTED_PIN`) で安全に注入可能です。PIN が記号／英字を含む場合は `UseEnhancedPin=1` を自動設定します。

## 入力 (CSV)
`bitlocker_list.csv` の主な列:
- `Enabled` … 1=対象 / 0=スキップ
- `TargetDrive` … 対象ドライブ（例: `C:`, `D:`）
- `EncryptionMethod` … `XtsAes128` / `XtsAes256` 等
- `UsedSpaceOnly` … TRUE/FALSE（使用領域のみ暗号化）
- `SkipHardwareTest` … TRUE/FALSE
- `AutoUnlock` … TRUE/FALSE（データドライブ向け）
- `Pin` … TpmAndPin 用 PIN（空欄=TPM-only）。**`ENC:` プレフィックス対応**
- `Description` / `Segment`

PIN 解決優先順: `$env:SELECTED_PIN` > `Pin` 列 > なし（TPM-only）

## 主要ステップ
[Enable]
1. TPM 状態確認（`TpmPresent` 必須、`TpmReady=false` は Warning）
2. `bitlocker_list.csv` 読み込み（Enabled=1）
3. PIN 解決＋ENC: ガード判定
4. ドライブ存在確認＋現状表示
5. 実行確認（AutoPilot 時は自動 Y）
6. 必要なら FVE レジストリポリシー（`UseAdvancedStartup=1` / `UseTPMPIN=1` / `UseEnhancedPin=1`）設定
7. エビデンス出力先決定 → ドライブごとに `Enable-BitLocker` 実行 → 回復キー取得 → エビデンス書き出し → AutoUnlock 適用

[Disable] CSV 読込 → 状態判定 → 実行確認 → `Disable-BitLocker`
[Await]   CSV 読込 → Encrypting/Decrypting 検出 → 実行確認 → 30 秒ポーリング → 30 分進行なしで個別タイムアウト

## 注意点・運用メモ
- **管理者権限必須**（3 スクリプト共通）
- 回復キーは **必ず Evidence に保存**（TpmPin 内容は記録しない、セキュリティ配慮）
- AutoUnlock はシステムドライブが FullyEncrypted になってから有効化される仕様のため、Enable 直後の AutoUnlock 適用は失敗しがち（Warning のみで継続）。Await 通過後の再実行で適用される
- FVE レジストリポリシーを書き換えるため、グループポリシー管理環境では運用方針との衝突を事前確認
- Profile 構成例: `... → bitlocker_config.ps1 → bitlocker_await.ps1 → ...`

## 検証
3 スクリプトいずれも `-Verified` フラグは未返却。理由は「BitLocker の状態遷移は本質的に非同期」であり、Enable/Disable 直後に Verified=true を返してしまうと「完了」を意味してしまい誤解を招くためです。完了判定は `bitlocker_await` の責務に分離されています。Enable は「回復キー取得 + エビデンス保存」をもって受理成功扱い、Disable は `Disable-BitLocker` 受理をもって成功扱い、Await は `FullyEncrypted` / `FullyDecrypted` で集計します。

冪等性: Enable は `ProtectionStatus=On` のドライブを Skip（PIN 追加が必要なら Tpm→TpmAndPin にアップグレード）、Disable は `FullyDecrypted` / `DecryptionInProgress` を Skip。


<!-- ============================================================ -->
# === modules/bloatware_export.md ===
<!-- ============================================================ -->

# bloatware_export (Standard)

**カテゴリ**: Applications
**メニュー名**: Bloatware Export
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（出力系モジュール。システム状態を変更しない）
**サブスクリプト**: なし

## 目的
PC にインストールされているデスクトップアプリケーションを `HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall`（64bit / 32bit 両ハイブ）から列挙し、CSV としてエクスポートするインベントリ採取モジュールです。出力 CSV は `bloatware_remove` モジュールの入力フォーマットに揃えてあるため、「OEM 出荷時の不要アプリを 1 度棚卸して、削除対象だけ Enabled=1 に編集する」という現場ワークフローを支援します。納品エビデンスとしての「出荷時アプリ一覧」記録としてもそのまま使えます。

## 入力 (CSV)
入力 CSV は **無し**。出力 CSV のスキーマ:
- `Enabled` … 全行 0 で出力（編集して 1 にする）
- `DisplayName` … レジストリから取得したアプリ名
- `MatchPattern` … 空欄出力（必要に応じて `McAfee*` 等のワイルドカードに編集）
- `Description` … 空欄出力（任意ラベル）
- `Segment` … 空欄出力（Segment フィルタ用）
- `Publisher` / `Version` … 参考情報（DisplayName だけでは特定しづらい時の手掛かり）

## 主要ステップ
1. （CSV 読み込みなし）
2. Pre-flight: `Test-AdminPrivilege`
3. スキャン対象（HKLM Uninstall 64bit/32bit）と出力先表示
4. 実行確認（AutoPilot 時は自動 Y）
5. レジストリ列挙 → CSV ファイルへ出力
6. 結果サマリ

## 注意点・運用メモ
- **管理者権限必須**（HKLM 読み取り権限のため）
- 出力先は `evidence\inventory\app_inventory_<日時>_<PC名>.csv`
- ファイル名の PC 名サフィックス優先順: `$env:SELECTED_NEW_PCNAME` > `$env:COMPUTERNAME`
- `evidence\inventory\` ディレクトリは自動作成
- `module.csv` の Enabled は **0**（メニューに常時出さない運用）。必要時だけ手動実行する想定
- ストアアプリ（AppX/MSIX）は対象外。デスクトップアプリ（クラシック Win32 アンインストールエントリ）のみ

## 検証
Post-Apply Verification は **不要**（出力系のためシステム状態を変更しない）。`-Verified` は未渡しで Verified 列は空欄になります。出力 CSV のファイル生成成否は `New-ModuleResult` の Status / Message で判定する設計です。CSV 内容そのものは「現状の HKLM Uninstall を反映した snapshot」であり、検証ではなくキャプチャの位置づけです。


<!-- ============================================================ -->
# === modules/bloatware_remove.md ===
<!-- ============================================================ -->

# bloatware_remove (Standard)

**カテゴリ**: Applications
**メニュー名**: Bloatware Remove
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（HKLM Uninstall を再検索して残存エントリを判定）
**サブスクリプト**: なし

## 目的
OEM 出荷時にプリインストールされている不要アプリ（McAfee 系 / Lenovo Now / Microsoft 365 体験版 等）を CSV ベースで一括アンインストールします。アンインストール情報（QuietUninstallString / MSI ProductCode / UninstallString）はレジストリから **実行時に動的取得** するため、端末ごとのバージョン差異やインストールパス差異に左右されないのが特長です。`MatchPattern` のワイルドカードで「`McAfee*` ＝ McAfee 系全部」を 1 行で書けるため、CSV メンテが現実的に保てます。

## 入力 (CSV)
`bloatware_list.csv` の主な列:
- `Enabled` … 1=削除対象 / 0=スキップ
- `DisplayName` … アプリ名（MatchPattern 省略時は完全一致で検索）
- `MatchPattern` … ワイルドカードパターン（例: `McAfee*`）。1 パターンが複数アプリにマッチした場合は全削除
- `Description` … 表示ラベル
- `Segment` … Segment フィルタ

## 主要ステップ
1. `bloatware_list.csv` 読み込み + レジストリ照合（HKLM Uninstall 64bit/32bit ハイブ）
2. Pre-flight: `Test-AdminPrivilege`
3. 検出結果の一覧表示（dry-run）
4. 実行確認（AutoPilot 時は自動 Y）
5. アンインストール実行（`QuietUninstallString` → `MSI ProductCode` → `UninstallString` の優先順）
5.5 **Post-Apply Verification**: `Find-RegistryUninstallEntry` で再スキャンし、`NoRemove=1` / `SystemComponent=1` を除く残存エントリが無いか検証
6. `New-BatchResult ... -Verified $verified` で結果返却

## 注意点・運用メモ
- **管理者権限必須**
- `NoRemove=1` または `SystemComponent=1` のアプリはレジストリ値から自動判定し、対象から除外（自動回復不能なシステムコンポーネントを誤削除しないため）
- アプリのアンインストールパスを CSV に書く必要なし（実行時にレジストリから取得）
- ストアアプリ（AppX）は対象外。同種需要は `storeapp_config` モジュールで対応
- 1 つの MatchPattern が複数アプリにマッチした場合、すべて削除されるため広めのワイルドカードは事前 `bloatware_export` で確認推奨

## 検証
Post-Apply Verification を **実装あり** とする数少ないモジュールの 1 つ。Step 5.5 で `Find-RegistryUninstallEntry` をもう一度走らせ、削除対象だった行が `NoRemove`/`SystemComponent` を除いて完全に消えているかを判定します。残存があれば `[VERIFY FAILED]` をカウントし、`-Verified $verified` フラグ付きで `New-BatchResult` を返却。Evidence Manager 側で「Verified 列で削除完了が突合できる」状態になります。


<!-- ============================================================ -->
# === modules/brightness_config.md ===
<!-- ============================================================ -->

# brightness_config (Standard)

**カテゴリ**: Display
**メニュー名**: Brightness Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`WmiMonitorBrightness.CurrentBrightness` を読み返し）
**サブスクリプト**: なし

## 目的
ノート PC のキッティング工程で「内蔵ディスプレイの初期輝度を一定値（例: 80%）に揃える」という標準作業を CSV 1 行で表現するモジュールです。WMI クラス `root\wmi\WmiMonitorBrightnessMethods` を介して即時反映するため、再起動を要しません。外付けモニターのみで内蔵ディスプレイを持たないデスクトップ PC では WMI が応答しない仕様のため、本モジュールは **Error ではなく Skipped で安全終了** する設計になっており、共通プロファイルでデスクトップ／ノートを混載しても運用しやすいようになっています。

## 入力 (CSV)
`brightness_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Brightness` … 輝度（0〜100）
- `Description` … 表示用ラベル
- `Segment` … Segment フィルタ

## 主要ステップ
1. CSV 読み込み（`Import-ModuleCsv -FilterEnabled`）
2. WMI 輝度制御の対応可否確認（未対応なら Skipped で終了）
3. 現在の輝度と変更先を表示（dry-run）
4. 実行確認（AutoPilot 時は自動 Y）
5. WMI メソッドで輝度変更
5.5 **Post-Apply Verification**: `WmiMonitorBrightness.CurrentBrightness` を読み返し、CSV 目標値と一致するか検証
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- ノート PC など内蔵ディスプレイを持つデバイスでのみ動作
- 外付けモニターのみのデスクトップ PC は自動 Skipped で安全終了（Error 扱いにしない）
- 設定値は OS の電源プラン UI における「100% / 80% / 50%」スライダ位置に直接対応
- 環境変数は使用しない

## 検証
Post-Apply Verification は **実装あり**。WMI の `CurrentBrightness` を再取得して CSV 目標値との完全一致を判定し、一致なら Verified=true、不一致なら false で `New-BatchResult` に `-Verified` フラグを返却します。WMI 経由の即時反映なので、書き込み直後の読み返しが事実上信頼できる検証になっています。


<!-- ============================================================ -->
# === modules/browser_addon_config.md ===
<!-- ============================================================ -->

# browser_addon_config (Standard)

**カテゴリ**: Applications
**メニュー名**: Browser Addon Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（HKLM ForceList を再検索）
**サブスクリプト**: なし

## 目的
Chrome / Edge の拡張機能を **グループポリシー（HKLM レジストリ）経由** で強制インストール対象に登録するモジュールです。書き込み先は Chrome が `HKLM\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist`、Edge が `HKLM\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist`。Chrome Web Store URL を貼り付けるだけで 32 文字 ID を自動抽出し、既存 Forcelist の最大インデックス +1 を計算して衝突なく追記する設計のため、CSV メンテが極めて軽量です。ブラウザ再起動または `gpupdate /force` で拡張がダウンロード／インストールされます。

## 入力 (CSV)
`browser_addon_list.csv` の主な列:
- `Enabled` … 1=登録 / 0=スキップ
- `Browser` … `Chrome` または `Edge`
- `ExtensionId` … 32 文字 ID 直接、または Chrome Web Store URL（自動抽出）
- `Description` … 表示ラベル
- `Segment` … Segment フィルタ

ExtensionId 解決は 3 段階フォールバック: (1) 32 文字 a-p の ID 直接 → (2) `/detail/` 以降から抽出 → (3) 文字列中の最初の 32 文字 a-p 系列を抽出。

## 主要ステップ
1. `browser_addon_list.csv` 読み込み
2. 前処理（Browser 値検証、`Resolve-ExtensionId`、不正は ERROR フラグ）
3. Dry-run 表示（`[Current]` / `[Change]` / `[ERROR]` で色分け、書き込み先パスも併記）
4. 実行確認（AutoPilot 時は自動 Y）
5. 設定適用ループ（既登録は Skip、新規はキー作成 + `Get-NextForcelistIndex` で次インデックスを採番し `<id>;https://clients2.google.com/service/update2/crx` 形式で書き込み）
5.5 **Post-Apply Verification**: 全対象を `Test-ExtensionInForcelist` で再検証し `[VERIFIED]` / `[VERIFY FAILED]`
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- **管理者権限必須**（HKLM 書き込み）
- ブラウザ自体は本モジュールではインストール／確認しない（レジストリのみ）
- 反映にはブラウザ再起動 or `gpupdate /force` が必要
- **同一 ID の二重登録は発生しない**（冪等性あり: `Test-ExtensionInForcelist` で判定）
- update_url は `clients2.google.com/service/update2/crx` 固定。プライベート CRX レジストリには非対応
- 既存 Forcelist に他拡張が登録されていてもそれらは保持し、追記のみ
- 削除機能は本モジュールにはない（`reg_hklm_delete` 等で対応）
- Edge 拡張も Chrome Web Store の update_url を使用（Microsoft 推奨動作）

## 検証
Post-Apply Verification は **実装あり**。Step 5.5 で `Test-ExtensionInForcelist` をもう一度実行し、CSV で対象とした ID 全件がレジストリに登録されていることを検証します。`-Verified $verified` で `New-BatchResult` に返却。ただし「ブラウザが実際に拡張をダウンロード／インストールしたか」までは本モジュールでは検証範囲外で、最終確認は `chrome://extensions` / `edge://extensions` で行う運用です。


<!-- ============================================================ -->
# === modules/builtin_admin_config.md ===
<!-- ============================================================ -->

# builtin_admin_config (Extended)

**カテゴリ**: User Management
**メニュー名**: Built-in Admin Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（パスワード値が `Get-LocalUser` で読み返せず Verification 全体を意図的に非実装）
**サブスクリプト**: なし

## 目的
ビルトイン Administrator アカウントのパスワード設定・有効化/無効化・パスワード無期限フラグ・説明文を一括設定するモジュール。
`Enabled` 列が「実行有効化」と「Enable/Disable」の両方を兼ねる特殊仕様で、`0` の行でも処理対象として読み込まれ Disable 操作が走る（このため `Import-ModuleCsv -FilterEnabled` ではなく手動取得しているわけではないが、CSV を 1 行のみ前提としている）。

## 入力 (CSV)
`builtin_admin.csv`:
- **Enabled**: 1=Enable+パスワード設定, 0=Disable
- **Password**: 設定パスワード（`ENC:` プレフィックス暗号化対応）
- **PasswordNeverExpires**: 1=無期限 / 0=期限あり
- **Description**: アカウント説明文
- **Segment**: Segment フィルタ（`Import-ModuleCsv` が暗黙参照）

## 主要ステップ
1. CSV 読み込み（先頭 1 行のみ使用）
2. Password 空欄チェック
3. `Get-LocalUser` で Administrator 存在確認
4. 設定内容（Enable/Disable, パスワード期限, Description）を画面表示
5. `Confirm-ModuleExecution` で実行確認
6. Enable/Disable → Set-LocalUser でパスワード → PasswordNeverExpires → Description の順に適用

## 注意点・運用メモ
- パスワード値は冪等性チェックなしで毎回 `Set-LocalUser` を呼ぶ（`Get-LocalUser` ではパスワードを取得できないため構造的に比較不可）
- 何度実行しても結果は同じだが、書き込みは毎回発生する
- Administrator アカウントが存在しない環境（一部のクリーンインストール直後）では Error 終了
- 管理者権限必須（`Set-LocalUser` の前提）

## 検証
パスワード読み返し不可のため Verification 全体を未実装。Enabled / PasswordExpires は技術的には検証可能だが、肝心のパスワード検証ができないため整合性を保つ目的で全体非対応。`-Verified` 未指定で Verified 列は空欄。


<!-- ============================================================ -->
# === modules/cert_config.md ===
<!-- ============================================================ -->

# cert_config (Standard)

**カテゴリ**: Security
**メニュー名**: Certificate Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（サムプリントによるストア再検索）
**サブスクリプト**: なし（証明書実体は `certs/` 配下に配置）

## 目的
証明書ファイル（`.pfx` / `.p12` / `.cer` / `.crt`）を Windows の証明書ストアにインポートするモジュールです。**明示指定モード** （CSV でストアスコープ／ストア名を直接指定）と **AutoRoute モード** （PFX 内の各証明書を BasicConstraints CA フラグで判定し、CA は `LocalMachine\Root`、クライアントは `LocalMachine\My` に自動振り分け）の 2 動作を持ちます。FileName は環境変数 + ワイルドカード対応のため、`%SELECTED_NEW_PCNAME%_*.p12` のように書けば `hostlist.csv` ベースで「PC ごとに異なるクライアント証明書」を自動配布する運用が可能です。PFX パスワードは ENC: プレフィックス暗号化対応。

## 入力 (CSV)
`cert_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `FileName` … `certs/` 内のファイル名（環境変数 / ワイルドカード可）
- `StoreScope` … `LocalMachine` / `CurrentUser`（AutoRoute=1 時は空欄可）
- `StoreName` … `My` / `Root` / `CA` / `TrustedPublisher` 等（AutoRoute=1 時は空欄可）
- `Password` … PFX パスワード（**`ENC:` プレフィックス対応**、`.cer`/`.crt` は空欄）
- `AutoRoute` … 1=自動振り分け / 0=明示指定
- `Overwrite` … 1=置換 / 0=既存スキップ
- `FriendlyName` … クライアント証明書のフレンドリ名
- `Description` / `Segment`

## 主要ステップ
1. `cert_list.csv` 読み込み
2. `certs/` ディレクトリ存在確認
3. 各証明書の状態確認 + 対象一覧表示（`[IMPORT]`/`[SKIP]`/`[REPLACE]`）
4. 実行確認（AutoPilot 時は自動 Y）
5. インポート（明示指定 or AutoRoute で振り分け）
5.5 **Post-Apply Verification**: `Test-CertificateInStore` で `X509Store.Certificates` をサムプリント検索
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- **管理者権限必須**（`LocalMachine` ストア書き込み。明示チェックなしで失敗時に検出）
- AutoRoute モードでは複数証明書（ルート CA + 中間 CA + クライアント）を含む PFX も適切に振り分け
- `%SELECTED_NEW_PCNAME%` + ワイルドカードで PC 別証明書を自動マッチ可能
- ENC: 復号は `Import-ModuleCsv` 側で自動実施（マスターパスフレーズ要）
- CSV の文字エンコーディングは Shift-JIS / UTF-8 BOM 付き両対応（fabriq 共通）
- 環境変数: `$env:SELECTED_NEW_PCNAME` / `$env:SELECTED_SEGMENT`

## 検証
Post-Apply Verification は **実装あり**。Step 5.5 で実際にインポート／既存維持した全証明書を `Test-CertificateInStore` でサムプリント検索し、ストア内に存在することを確認します。AutoRoute モードでは PFX 内の各証明書を個別に検証。全件成功で Verified=true、1 件でも失敗すれば false。`-Verified` フラグ付きで `New-BatchResult` に返却するため、Evidence Manager 側で「証明書配布完了」を突合できます。


<!-- ============================================================ -->
# === modules/copyfile_config.md ===
<!-- ============================================================ -->

# copyfile_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: File Copy Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 不可（誤 PASS リスクのため意図的に非実装）
**サブスクリプト**: なし（コピー元実体は `source/` 配下に配置）

## 目的
`source/` ディレクトリに配置したファイル／フォルダを、CSV に記述された宛先（`%USERPROFILE%`、`%LOCALAPPDATA%`、`C:\Program Files\...` 等）にコピーするモジュールです。フォルダ指定時は `Copy-Item -Recurse -Force` で中身ごと再帰コピーします。コピー後に **Mark-of-the-Web（Zone.Identifier ADS）を自動解除** するため、ネットワーク経由で受け取ったテンプレート／設定ファイルでも「ブロック解除済み」状態で配置されるのが運用上の利点です。`Expand-UserEnvironmentVariables` により昇格セッションでもログオンユーザーの `%USERPROFILE%` を解決します。

## 入力 (CSV)
`copy_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `FileName` … `source/` 直下のファイル名 or フォルダ名（フォルダなら再帰）
- `DestPath` … コピー先ディレクトリ（環境変数展開対応: `%USERPROFILE%` / `%LOCALAPPDATA%` / `%APPDATA%` / `%ProgramData%` / `%SystemRoot%` 等）
- `Overwrite` … 1=上書き / 0=既存スキップ
- `Description` / `Segment`

## 主要ステップ
1. `copy_list.csv` 読み込み
2. `source/` ディレクトリ存在確認（無ければ Error）
3. （前処理）DestPath 環境変数展開
4. Dry-run 表示（`[Missing]` / `[Current]` / `[Overwrite]` / `[Copy]`）
5. 実行確認（AutoPilot 時は自動 Y）
6. コピー実行（DestPath 自動作成 → `Copy-Item -Recurse -Force` → `Remove-ZoneIdentifier` で MoTW 解除）
7. `New-BatchResult` で結果集計

## 注意点・運用メモ
- 管理者権限は **状況依存**: `%USERPROFILE%` 等のユーザー領域なら不要、`C:\Program Files` 等の保護領域なら必須
- フォルダ指定時は `-Recurse -Force` で丸ごと上書き。個別ファイル選別は不可
- UNC パス（`\\server\share`）は認証次第で失敗。事前 `net use` 推奨
- `Overwrite=0` 指定は既存ファイル保護で擬似冪等性あり（タイムスタンプ更新も発生しない）

## 検証
Post-Apply Verification は **意図的に非実装**（CLAUDE.md 記載の除外モジュール）。理由は、`Copy-Item` 成功後に `Test-Path` で存在を見るだけだと「コピーしたつもりが内容が古い」「権限が前のまま」「実体が古いコピー」といったケースで誤 VERIFIED となるリスクが大きいためです。真の検証には内容ハッシュ比較・ACL 比較・MoTW 二重確認が必要ですが、配置先・用途が多様で一律実装が困難なため、安易な検証で誤 PASS を出さない方針を選択しています。代替として、コピー直後のログで `[Copy]` / `[Overwrite]` / `[Skip]` / Error を明確に記録し、Mark-of-the-Web は `Remove-ZoneIdentifier` で必ず解除しています。`-Verified` は未渡しで Verified 列は空欄。厳密検証は運用側で `Get-FileHash` / `icacls` を別途実行します。


<!-- ============================================================ -->
# === modules/default_app_config.md ===
<!-- ============================================================ -->

# default_app_config (Standard)

**カテゴリ**: System
**メニュー名**: Export App Associations / Default App Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（新規ユーザープロファイル作成時に反映される仕様のため検証困難）
**サブスクリプト**:
- `default_app_config.ps1` … 本番キッティング用（`Dism /Import-DefaultAppAssociations`）
- `export_app_associations.ps1` … マスター準備用（`Dism /Export-DefaultAppAssociations`）

## 目的
Windows 10/11 で「PDF を Acrobat Reader で開く」「ブラウザは Chrome を既定にする」など、**既定アプリの関連付け** を ProgId ハッシュ保護に対応した安全な方法で配布するモジュールです。マスター PC で手動設定 → エクスポートで XML 化（`Export-DefaultAppAssociations`）、本番 PC で XML をインポート（`Import-DefaultAppAssociations`）する 2 段構成。インポートは「**新規ユーザープロファイル作成時** に適用」される Windows の仕様に基づいているため、キッティングで sysprep 前に書いておくのが定石です。Segment 列で「営業部用 / 技術部用」など部署別の関連付けセットを切り替え可能。

## 入力 (CSV)
`default_app_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `XmlFile` … `xml/` 配下の XML ファイル名
- `Description` … 表示ラベル
- `Segment` … セグメント名（部署／拠点別の使い分け）

## 主要ステップ（Import 側 = `default_app_config.ps1`）
1. `default_app_list.csv` 読み込み
2. Pre-flight（DISM 利用可否、`xml/` ディレクトリ）
3. Dry-run 表示
4. 実行確認（AutoPilot 時は自動 Y）
5. `Dism /Online /Import-DefaultAppAssociations:"<XML>"` を順次実行
6. 結果集計（`New-BatchResult`）

Export 側（`export_app_associations.ps1`）は同じスケルトンで `/Export-DefaultAppAssociations` を実行し、`xml/` に上書き出力。

## 注意点・運用メモ
- **管理者権限必須**（`Dism /Online` は管理者専用）
- インポートは **新規ユーザープロファイル作成時に反映**。既存ユーザーには即時反映されない（sysprep 前に適用するのが標準）
- エクスポートは現在ログオンユーザーの関連付けを取得。既存 XML は毎回上書き
- 冪等性は **非対称**: Export は毎回 DISM 実行（スナップショット用途）、Import は XML 不在で Skip / 存在で毎回 DISM 実行（差分チェックなし）
- 環境変数は使用しない。XML パスは `$PSScriptRoot\xml` 配下で解決

## 検証
Post-Apply Verification は **未実装**。`Dism /Import-DefaultAppAssociations` の `ExitCode` のみで成否判定し、`-Verified` は未渡し（Verified 列は空欄）。理由は、関連付けの実反映が「次回新規ユーザープロファイル作成時」という遅延型仕様のため、適用直後にレジストリを読んでも何も変わっておらず検証が事実上不可能なためです。エビデンスとしては XML ファイルの存在と DISM ExitCode で代替し、実反映は新規ユーザーログオン後に手動確認する運用です。


<!-- ============================================================ -->
# === modules/desktop_icon_config.md ===
<!-- ============================================================ -->

# desktop_icon_config (Extended)

**カテゴリ**: Desktop
**メニュー名**: Desktop Icon Backup / Desktop Icon Restore（2 メニュー）
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（Backup は ExitCode + ファイルサイズで成否、Restore はサインアウト後反映のため即時検証無意味）
**サブスクリプト**: `desktop_icon_backup.ps1`（エクスポート役）+ `desktop_icon_restore.ps1`（インポート役）。`backup/` ディレクトリにタイムスタンプ付き `.reg` を蓄積

## 目的
デスクトップアイコンの配置情報を保持するレジストリキー `HKCU\Software\Microsoft\Windows\Shell\Bags\1\Desktop` を `.reg` ファイルにバックアップ／復元するペアモジュール。
SYSTEM 起動セッションでも動作するように `Resolve-HkcuRoot` で対象ハイブを解決し、エクスポート時は `HKEY_USERS\<SID>` を `HKEY_CURRENT_USER` に正規化してポータブルな `.reg` を生成、インポート時は逆方向の書き換えを TEMP に作って `reg import` する仕組み。

## 入力 (CSV)
なし（対象レジストリパスはハードコード）。

## 主要ステップ（Backup）
1. `Resolve-HkcuRoot` で書き込み先ハイブ（HKCU or HKU\<SID>）を決定
2. 対象キーの存在確認 + 状態表示
3. `Confirm-ModuleExecution` で実行確認
4. `backup/` を必要なら新規作成
5. `reg.exe export` でタイムスタンプ付きファイル名 (`DesktopIcons_yyyyMMdd_HHmmss.reg`) に出力
6. HKCU リダイレクト時は出力 `.reg` 内の `HKEY_USERS\<SID>` を `HKEY_CURRENT_USER` に置換して可搬化

## 主要ステップ（Restore）
1. `backup/` から `DesktopIcons_*.reg` を新しい順に列挙（最大 5 件まで一覧表示）
2. 最新ファイルをターゲットとして表示
3. `Confirm-ModuleExecution` で実行確認
4. HKCU リダイレクト時は `.reg` の `HKEY_CURRENT_USER` を `HKEY_USERS\<SID>` に置換した一時ファイルを TEMP に作成
5. `reg.exe import` で書き戻し
6. 一時ファイルクリーンアップ（`finally`）

## 注意点・運用メモ
- バックアップ前にアイコンを希望の配置に並べておく必要あり
- 復元の反映にはサインアウトまたは再起動が必要
- 管理者権限必須（HKU 操作のため）
- 過去のバックアップは `backup/` 内に蓄積され続けるため、不要なら手動削除

## 検証
両スクリプトとも `reg.exe` の ExitCode 判定のみ。`.reg` の中身検証や、Restore 後のキー値読み返しは行わない。`-Verified` 未渡しで履歴の Verified 列は空欄。


<!-- ============================================================ -->
# === modules/directory_cleaner.md ===
<!-- ============================================================ -->

# directory_cleaner (Extended)

**カテゴリ**: Maintenance
**メニュー名**: Directory Cleaner
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（contents モードで削除直後に新規一時ファイルが生成され得るため false PASS/FAIL を回避）
**サブスクリプト**: なし

## 目的
CSV で指定したディレクトリの「中身のみ削除（contents）」または「ディレクトリごと削除（directory）」を実行するクリーンアップモジュール。
最大の特徴はハードコードされた**禁止パスホワイトリスト**による多重ガード。`C:\Windows`, `C:\Program Files`, `%USERPROFILE%`, fabriq リポジトリルートなどシステム重要パスを 3 種類のチェック（完全一致 / 親子関係 / セグメント数 < 3）で自動ブロックする。表示ステップと実行ステップの両方で二重チェックされ、ヒットすると `[BLOCKED]` でスキップ。

## 入力 (CSV)
`clean_list.csv`:
- **Enabled**: 有効フラグ
- **TargetPath**: 対象ディレクトリ（`%LOCALAPPDATA%` 等の環境変数展開対応）
- **Mode**: `contents`（中身のみ削除、フォルダは残す） / `directory`（フォルダごと削除）
- **Description**: 表示用説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. CSV 読み込み + `Expand-UserEnvironmentVariables` で環境変数展開
2. Mode 値検証（`contents`/`directory` 以外は Error）
3. ドライラン表示（`[BLOCKED]` / `[NOT FOUND]` / `[DELETE]` を色分け）
4. `Confirm-ModuleExecution`
5. 削除ループ（実行時にも `Test-ForbiddenPath` を再評価する double-gate / 不在 / 既に空 → Skip）
6. `New-BatchResult` で集計

## 注意点・運用メモ
- 禁止パスチェックは 3 種類（A: 完全一致、B: 対象が禁止パスの親、C: セグメント数 3 未満）
- 削除中にロックされたファイルは `SilentlyContinue` で個別スキップし、件数を集計表示（contents モードでは部分削除を Warning 扱い）
- Past incident: 2026-04-25 にユーザーのデスクトップを Join-Path 失敗で破壊した事故への対策として禁止パスガードが入っている経緯あり

## 検証
contents モードでは削除直後に OS が新規一時ファイルを生成する性質があり、読み返し検証は false FAIL/PASS リスクが高いため意図的に未実装。`-Verified` 未渡しで履歴の Verified 列は空欄。


<!-- ============================================================ -->
# === modules/display_config.md ===
<!-- ============================================================ -->

# display_config (Extended)

**カテゴリ**: Display
**メニュー名**: Display Resolution Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（適用反映が再起動後のため即時読み返し検証は無意味、ただし冪等性チェックは実装あり）
**サブスクリプト**: なし

## 目的
ディスプレイ解像度をレジストリ `HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration\<HardwareID>` 配下の `PrimSurfSize.cx/cy` 書き換えで設定するモジュール。
standard 側の `resolution_api_config`（DisplayConfig API による即時反映、単一ディスプレイのみ）と棲み分けて、こちらは**複数モニター・HardwareID 指定・再起動後反映**を担う extended 版。

## 入力 (CSV)
`display_list.csv`:
- **Enabled**: 有効フラグ
- **HardwareID**: ディスプレイ ID（`AUTO`=自動検出 / 具体的 ID は前方一致 / 空欄=対話選択）
- **Width** / **Height**: ピクセル
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み + Width/Height の正数検証
3. 対象ディスプレイ解決（AUTO=単一なら自動 / 複数 or 0 件は Interactive にフォールバック / 具体的 ID は前方一致）
4. ドライラン表示（`[Current]` / `[Change]` / `[Interactive]` マーカー、現在解像度併記）
5. `Confirm-ModuleExecution`
6. 各キーの最初のサブキーに対し冪等性チェック → `PrimSurfSize.cx/cy` を DWORD で書き込み
7. `New-BatchResult` 集計（成功 1 件以上で「再起動が必要」警告）

## 注意点・運用メモ
- 反映には**再起動が必須**
- AutoPilot 運用時は HardwareID を空欄/AUTO（複数モニター環境）にすると Interactive にフォールバックしてブロックされるため、具体的 ID 指定推奨
- 同じ Description で複数キーがマッチした場合はマッチした全キーに書き込む

## 検証
冪等性チェック（適用前の `PrimSurfSize.cx/cy` 読み出し）は実装ありだが、再起動が必要な性質上 Post-Apply Verification は省略。`-Verified` 未渡しで履歴の Verified 列は空欄。


<!-- ============================================================ -->
# === modules/domain_join.md ===
<!-- ============================================================ -->

# domain_join (Standard)

**カテゴリ**: Network
**メニュー名**: Domain Join
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（再起動後に反映されるため、再起動前の検証では情報量が足りない）
**サブスクリプト**: なし

## 目的
PC を Active Directory ドメインに参加させるモジュールです。`Add-Computer` コマンドレットを使用し、参加に失敗した場合は **GUI ダイアログでエラー内容を表示してリトライループ** に入る設計のため、現場担当者がパスワード入力ミスなどに対処しやすくなっています。AutoPilot 運用時は確認スキップで自動参加させますが、リトライダイアログが立ち上がる仕様上、Profile 側で `ErrorMode=retry` と組み合わせる運用が想定されます。

## 入力 (CSV)
`domain.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `domain` … ドメイン名（例: `example.local`）
- `user` … ドメイン参加用アカウント
- `pass` … パスワード（**`ENC:` プレフィックス暗号化対応**）
- `dns` … DNS サーバ IP（接続確認 Ping 用）
- `Description` / `Segment`

## 主要ステップ
1. `domain.csv` から有効エントリ読み込み
2. DNS サーバへの Ping 接続確認
3. 実行確認（AutoPilot 時は自動 Y）
4. `Add-Computer` でドメイン参加（失敗時は GUI ダイアログでリトライループ。`adminstop` 入力で中断）
5. `New-ModuleResult` で結果返却

## 注意点・運用メモ
- **管理者権限必須**（`Add-Computer` 実行のため）
- DNS サーバへのネットワーク接続が必須（Ping 失敗で Error 検出）
- ENC: 暗号化パスワードを使う場合は Fabriq 起動時のマスターパスフレーズ入力が必要
- リトライダイアログは無人運用に向かない。AutoPilot 構成時は Profile 側 `ErrorMode=retry` と組み合わせ
- 反映は **再起動後**（Windows 仕様）。本モジュールは再起動はトリガーしない（後続 `restart_config` で集約再起動）

## 検証
Post-Apply Verification は **未実装**。理由は、ドメイン参加状態は `Add-Computer` 実行直後ではなく **次回再起動後** に `(Get-WmiObject Win32_ComputerSystem).PartOfDomain = $true` として反映される Windows 仕様のため、再起動前のメモリ上で検証しても判定が曖昧になるためです。`-Verified` は未渡しで Verified 列は空欄。`Add-Computer` の成否とリトライダイアログでの到達可否で結果判定し、再起動後の確認は手動または `evidence_config` の収集レポートで行う運用です。


<!-- ============================================================ -->
# === modules/dpi_api_config.md ===
<!-- ============================================================ -->

# dpi_api_config (Standard)

**カテゴリ**: Display
**メニュー名**: DPI Scaling Config (Live)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（適用前の冪等性チェックには `GetCurrentDpi` を使用するが、適用後の読み返し検証は省略）
**サブスクリプト**: なし

## 目的
ディスプレイの **DPI スケーリング（拡大率）を Windows API 経由で即時反映** させるモジュールです。レジストリ書き込み + サインアウト方式（`resolution_api_config` 系統）と異なり、`NativeDpiHelper::SetDpi` を呼び出すライブ方式のため、再起動／サインアウトなしで即座に変更が反映されるのが特徴です。「Live」というメニュー名はこのライブ反映を表しています。複数モニター環境では `MonitorIndex`（0=プライマリ）で個別に指定できます。

## 入力 (CSV)
`dpi_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `MonitorIndex` … 0=プライマリ、1=セカンダリ、…
- `ScalePercent` … 拡大率（100 / 125 / 150 / 175 / 200 等、Windows がサポートする刻み）
- `Description` / `Segment`

## 主要ステップ
1. `dpi_list.csv` 読み込み（Enabled=1）
2. 対象モニターと拡大率を表示
3. （冪等性チェック）`NativeDpiHelper::GetCurrentDpi` で現在値読み出し、目標と一致なら Skip
4. 実行確認（AutoPilot 時は自動 Y）
5. Windows API 経由で DPI を変更（即時反映）

## 注意点・運用メモ
- ノート PC の高 DPI 表示に合わせて 150% を設定するなど、現場では頻出のオペレーション
- レジストリ書き込み方式（`resolution_api_config` の dpi_reg 系）と違い、再起動／サインアウト不要
- 環境変数は使用しない
- Windows のサポート刻みから外れた `ScalePercent` を指定すると API 側で丸め or 失敗する可能性あり

## 検証
Post-Apply Verification は **未実装**。`NativeDpiHelper::GetCurrentDpi` は **適用前の冪等性チェック**（既に目標値ならスキップ）にのみ使用し、適用後の読み返し検証は行いません。理由は、API 反映直後と OS 内部のメトリック更新の間にラグがあり、即座の読み返しは false negative を出しやすいためです。`-Verified` は未渡しで Verified 列は空欄。実反映は次回ログオン後の見た目で確認する運用です。


<!-- ============================================================ -->
# === modules/dpi_config.md ===
<!-- ============================================================ -->

# dpi_config (Extended)

**カテゴリ**: Display
**メニュー名**: Display DPI Scaling Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（書き込みごとに HKCU/HIVE 値の冪等性チェック実装あり、ただし -Verified 未渡し）
**サブスクリプト**: なし（C# `DpiScaleResolver` クラスを `Add-Type` で動的コンパイル）

## 目的
モニターごとの DPI スケーリング（拡大率 100/125/150/175/200%）を、HKCU と Default プロファイル ntuser.dat（HKU\Hive）両方に同時書き込みするモジュール。
standard の `dpi_api_config`（即時反映・単一ディスプレイのみ）と異なり、こちらは**複数モニター・HardwareID 指定・新規ユーザー継承**まで対応。
DisplayConfig API（`DisplayConfigGetDeviceInfo` 等）を P/Invoke で叩く C# クラス `DpiScaleResolver` を内包し、モニターごとの推奨 DPI（recommended）からの相対値で `DpiValue` を算出する点が技術的特徴。

## 入力 (CSV)
`dpi_list.csv`:
- **Enabled**: 有効フラグ
- **HardwareID**: `AUTO` / 具体的 ID（前方一致） / 空欄（Interactive Display）
- **ScalePercent**: 100/125/150/175/200 のいずれか / 空欄や 0（Interactive Scale）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` + `Resolve-HkcuRoot` で HKCU 解決（昇格セッション対応）
2. C# `DpiScaleResolver` を `Add-Type` で読み込み、各モニターの推奨 DPI マップを構築
3. CSV 読み込み + ScalePercent 値の検証（サポート値以外はスキップ、空欄は Interactive Scale 扱い）
4. 対象解決（PerMonitorSettings 検索 → なければ GraphicsDrivers\Configuration へフォールバック → AUTO 複数や空欄は Interactive へ）
5. ドライラン表示（`[Current]`/`[Change]`/`[Interactive]` 色分け、推奨値 Recommended 表示）
6. `Confirm-ModuleExecution`
7. `reg load HKU\Hive C:\Users\Default\ntuser.dat` で Default ハイブをロード
8. 書き込みループ（ScalePercent → DpiValue 換算 → HKCU と HIVE 各々で冪等性チェック → 書き込み）
9. `reg unload`（失敗時は GC 強制 + sleep してリトライ）
10. `New-BatchResult` 集計

## 注意点・運用メモ
- 反映にはサインアウト/再起動が必要
- Default プロファイル書き込みのため、本モジュールは Profile 全体（hostname → 各種設定 → 新規ユーザー作成）の一環として使う想定
- AutoPilot で完全自動にしたい場合は HardwareID と ScalePercent の両方を CSV で具体指定すること（Interactive にフォールバックすると入力待ちでブロック）
- 推奨 DPI 取得失敗時はフォールバック値 150% を使用

## 検証
HKCU と HIVE 双方で書き込み前に現在 `DpiValue` を取得して冪等判定（一致時 Skip）はあるが、Post-Apply Verification は未実装。`-Verified` 未渡しで履歴 Verified 列は空欄。手動検証は `Get-ItemProperty 'HKCU:\Control Panel\Desktop\PerMonitorSettings\<key>' -Name DpiValue`。


<!-- ============================================================ -->
# === modules/driver_config.md ===
<!-- ============================================================ -->

# driver_config (Standard)

**カテゴリ**: System
**メニュー名**: Driver Export / Driver Import
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（限定的な検査として `*.inf` カウントと `pnputil` ExitCode は判定）
**サブスクリプト**:
- `driver_export_config.ps1` … `Export-WindowsDriver -Online` で現 OS のサードパーティドライバを `driver/{model}/` に書き出し
- `driver_import_config.ps1` … `pnputil /add-driver "*.inf" /subdirs /install` で対象 PC に展開

## 目的
キッティング工程で「マスター PC からドライバ一式をエクスポート → 同モデルの量産機にインポート」という典型ワークフローを CSV 1 行で表現するモジュールです。Export では現 OS にインストールされたサードパーティドライバ（OS 標準ドライバを除く）を `driver/{モデル名}/` 配下に展開し、Import では `pnputil` でそれを再投入します。`model` 列を空欄にすると **`Win32_ComputerSystem.Model` から自動取得** するため、ホスト名ベースではなく機種ベースで自動振り分けが可能です（例: ThinkPad X1 → `ThinkPad_X1` フォルダ）。

## 入力 (CSV)
`driver.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Id` … 識別連番
- `model` … ドライバフォルダ名の明示指定（空欄ならホストの `Win32_ComputerSystem.Model` を自動採用、サニタイズ: 空白→`_`、`\/:*?"<>|` 除去、80 文字上限）
- `Segment` … Segment フィルタ

## 主要ステップ
（Export / Import で同形）
1. `driver.csv` 読み込み
2. Pre-flight: 管理者権限確認 / `driver/` ディレクトリ存在確認
3. Dry-run 表示
4. 実行確認（AutoPilot 時は自動 Y）
5. 適用ループ（Export: `Export-WindowsDriver -Online -Destination ...`、Import: `pnputil /add-driver`）
6. 結果集計

## 注意点・運用メモ
- **管理者権限必須**（両スクリプトとも）
- Export はサードパーティドライバのみ対象（OS 標準ドライバは含まれない）
- Export 時、既存の `driver/{model}/` フォルダは中身がクリアされてから再エクスポート（差分追記ではなく完全置換）
- Import 後に再起動が必要な場合あり（`pnputil` 終了コード 3010）
- Export 時はドライバ容量分のディスク空き容量が必要

## 検証
Post-Apply Verification は **限定的**。`-Verified` フラグは渡しません。
- Export: 出力先フォルダの `*.inf` 数をカウント表示し、0 件ならエラー扱い
- Import: 実行前後の `*.inf` カウント参照 + `pnputil` ExitCode（0/3010=Success）

OS に組み込まれたドライバのバージョン比較等の厳密検証は行いません。必要に応じ `pnputil /enum-drivers` で手動確認する運用です。


<!-- ============================================================ -->
# === modules/evidence_config.md ===
<!-- ============================================================ -->

# evidence_config (Standard)

**カテゴリ**: Evidence
**メニュー名**: Collect Evidence
**VERSION**: 1.6.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 概念的に不適用（読み出し系のためシステム状態を変更しない）
**サブスクリプト**: なし

## 目的
PC のシステム情報を **22+ セクション** にわたって収集し、テキスト／CSV／HTML のエビデンスファイル一式を `evidence/pc_information/<日付>_<UID>_<PC名>/` に書き出す、fabriq の納品エビデンス採取の中核モジュールです。受入検査・官公庁監査・トラブル時の証跡として機能し、収集完了時に **`manifest.json`** を生成して各セクションの `id / title / files / status / reason / elapsedMs` を機械可読サマリ化します。これにより外部ツール（fabriq_evidence_manager 等）が manifest を起点にエビデンスを統合できる契約構造になっています。シリアル番号は 4 ソース横断 + canonical 選定 + OEM 無効値（"Default string" 等）の拒否までやる多重ソース戦略。

## 入力 (CSV)
**設定 CSV なし**。すべてのセクションが固定で実行されます。

## 収集セクション（抜粋）
1. システム基本情報 / 2. ローカル管理者 / 3. ネットワーク / 4. プリンター / 5. BitLocker / 6. MAC アドレス
7. PC シリアル番号（多重ソース + canonical 選定）／ 10. シリアル番号別ファイル
8. デスクトップアプリ / ストアアプリ CSV
9. ファイアウォール（プロファイル + ルール）CSV
10/11. Windows 機能 / Server Roles & Features CSV
20. System TEMP テキストログバックアップ（セーフティネット）
21. Windows ライセンス（SoftwareLicensingProduct + slmgr /dlv）
22. Office ライセンス（OSPP + vNext per-user スキャン + 自動解釈 verdict）
23. Security Baseline（TPM / Secure Boot / VBS / HVCI / Credential Guard / LSA Protection / BIOS）
24. Group Policy Report（`gpresult /h` HTML + サマリ TXT）
25. Certificates CSV（4 ストア統合、HasPrivateKey フラグのみ）
26. Battery Report（`powercfg /batteryreport` HTML、受入検査での容量契約エビデンス）

## 主要ステップ
1. 出力先パス決定（`$global:FabriqEvidenceBasePath` 有無で統一パスモード／レガシーモード分岐）
2. 実行確認（AutoPilot 時は自動 Y）
3. 各セクションを順次収集
4. 統合ファイル `_ALL_<PC名>_Log.txt` + 個別ファイル + CSV/HTML 出力
5. `manifest.json` 生成（`schemaVersion=1`、再実行時は `manifest.json.bak` に rotate）

## 注意点・運用メモ
- **管理者権限必須**（BitLocker / Firewall ルール / Defender 状態など admin 限定情報を含む）
- §22d の自動解釈は M365 sub の OSPP「飾りキー」問題（OSPP は常に Grace 表示になる）を内部で吸収し、Manifest の `Partial` / `Failed` 区別を出力
- §23 の各 probe は inner try/catch で個別退避（1 つ失敗してもセクションは Success）
- §24 の `gpresult` ユーザー側 RSoP は実行ユーザー（kitting admin01 等）視点であり、最終エンドユーザー視点ではない点に注意
- §26 はバッテリ非搭載で Skipped 完結
- 環境変数: `$env:SELECTED_NEW_PCNAME` / `$env:COMPUTERNAME`、グローバル `$global:FabriqUniqueId` / `$global:FabriqEvidenceBasePath`

## 検証
Post-Apply Verification は **概念的に不適用**。本モジュールは「読み出し系」であり、システム状態を変更しないため「適用後の読み返し検証」が論理的に存在しません。`-Verified` は未渡しで Verified 列は空欄。代わりに `manifest.json` のセクション単位 `status` フィールド（Success / Skipped / Failed / Partial の 4 値）が品質情報を担います。`kernel/EVIDENCE_MANIFEST.md` で公式契約として `schemaVersion=1` が定義されており、外部 evidence consumer はこの manifest を起点にエビデンスをパースする設計です。


<!-- ============================================================ -->
# === modules/fabriq_app_launcher.md ===
<!-- ============================================================ -->

# fabriq_app_launcher (Standard)

**カテゴリ**: Applications
**メニュー名**: Fabriq App Launcher
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（起動系のため概念的に不適用）
**サブスクリプト**: なし

## 目的
fabriq 内蔵アプリ（`apps/` ディレクトリ配下の `winget_gui` / `storeapp_editor` / `local_user_setup` 等の WPF/PowerShell GUI ツール群）を **Profile 実行中に挟んで起動** するためのブリッジモジュールです。「Profile の自動化フローの中で、ある時点だけ手動 GUI 操作を要求したい」というユースケース（例: Local User の対話設定、Store App の選別チェック、Winget の対話インストール）に対応します。`Wait=1` 指定でアプリ終了を待機して次モジュールへ進むため、プロファイル全体の進行制御も可能です。

## 入力 (CSV)
`target_apps.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `AppName` … `apps/` 内のアプリディレクトリ名（規約: `apps/{Name}/{Name}.ps1`）
- `Wait` … 1=アプリ終了まで待機 / 0=起動後すぐ次へ
- `Description` / `Segment`

## 主要ステップ
1. `target_apps.csv` 読み込み
2. Pre-flight: `apps/` ディレクトリ存在確認、各 `apps/{Name}/{Name}.ps1` 実在確認
3. Dry-run 表示
4. 実行確認（AutoPilot 時は自動 Y）
5. 各アプリを別プロセスとして起動（`Wait=1` なら終了待機）
6. 結果集計

## 注意点・運用メモ
- `apps/` ディレクトリはモジュール位置の 3 階層上（fabriq ルート直下）
- 各アプリは `apps/{Name}/{Name}.ps1` の命名規則に従う必要あり
- `powershell.exe` が PATH 解決可能であること（標準 Windows 環境では常時 OK）
- `Wait=1` のアプリは閉じられるまで Profile 進行をブロックするため、無人運用には不向き（手動立ち会い前提）
- Profile での典型用途: 「キッティングの最終局面で `local_user_setup` を起動して、現場担当者が個別ユーザー登録」など

## 検証
Post-Apply Verification は **概念的に不適用**。本モジュールは GUI アプリの起動成否（プロセス起動が成功したか、Wait 時は ExitCode）のみを扱うため、「設定が反映されたかの検証」概念がありません。`-Verified` は未渡しで Verified 列は空欄。アプリ自体の処理結果（例: `local_user_setup` で実際にユーザーが作成されたか）は、後続の `evidence_config` のローカル管理者一覧収集などで間接的に確認する運用です。


<!-- ============================================================ -->
# === modules/file_delete.md ===
<!-- ============================================================ -->

# file_delete (Standard)

**カテゴリ**: Maintenance
**メニュー名**: File Delete
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`Test-Path` で削除後の非存在を検証）
**サブスクリプト**: なし

## 目的
キッティング工程で「マスターイメージから引き継いだ不要ファイル」「テスト用一時ファイル」「アプリケーションのキャッシュフォルダ」などを CSV ベースで一括削除するモジュールです。`%TEMP%` / `%LOCALAPPDATA%` などの環境変数展開に対応し、`IfNotFound=Skip` / `IfNotFound=Error` で「存在しなくても OK」「絶対あるべきもの」を行単位に区別できる設計のため、冪等な再実行と「必須ファイルの取りこぼし検出」の両立が可能です。

## 入力 (CSV)
`delete_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Description` … 表示ラベル
- `TargetPath` … 削除対象パス（環境変数展開対応）
- `IfNotFound` … `Skip`（既定: 不在でも正常終了）/ `Error`（不在ならエラーとして記録）
- `Segment` … Segment フィルタ

## 主要ステップ
1. `delete_list.csv` 読み込み + 各パスの存在状況表示
2. 実行確認（AutoPilot 時は自動 Y）
3. 各対象を順次削除（フォルダは再帰削除）
4. **Post-Apply Verification**: 削除後に `Test-Path` で対象パスの非存在を全件検証
5. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- 管理者権限は **状況依存**（`Program Files` 等の保護領域なら必須、ユーザー領域なら不要）
- フォルダ指定時は再帰削除のため、誤って親フォルダを書くと連鎖的に大量削除される。CSV メンテ時は注意
- `IfNotFound=Skip` を使えば「テストで一時ファイルを置いた → 本番マスターでは存在しない」状況でも冪等
- `IfNotFound=Error` は「絶対消すべきファイル（個人情報の残骸など）」の検出に使う

## 検証
Post-Apply Verification は **実装あり**。Step 4 で `Test-Path` を使って対象パスの非存在を検証し、全件 PASS のとき `-Verified $true`、1 件でも残存していれば `$false` を `New-BatchResult` に渡して返却。実行履歴の Verified 列に結果が記録され、Evidence Manager 側で削除完了が突合可能です。`copyfile_config` の Verification 非実装と対照的に、削除側は「不在の検証」が `Test-Path` で確実に判定できるため実装している、という設計判断です。


<!-- ============================================================ -->
# === modules/firewall_config.md ===
<!-- ============================================================ -->

# firewall_config (Standard)

**カテゴリ**: Security
**メニュー名**: Firewall Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`Get-NetFirewallProfile` で読み返し）
**サブスクリプト**: なし

## 目的
Windows ファイアウォールの **3 プロファイル（Domain / Private / Public）を個別に有効／無効** に設定するモジュールです。Profile レベルの粗い粒度に専念しており、ルール単位の管理は `firewall_rule_config`（全体 backup/restore）と `firewall_rule_make_config`（個別作成）に分離されています。CSV を使う標準モード以外に、**Manual mode**（CSV が空ならインタラクティブメニューでトグル選択）と **Legacy CSV mode**（旧形式の `status` 列のみ）の 3 形態を統一スクリプトで吸収するため、レガシー資産も保護されます。

## 入力 (CSV)
`firewall_list.csv` の主な列:
- `Enabled` … 1=適用 / 0=スキップ
- `Profile` … `Domain` / `Private` / `Public`
- `Status` … `on`（有効）/ `off`（無効）
- `Description` / `Segment`

## 主要ステップ
1. 現在の 3 プロファイル状態取得＋表示
2. CSV または Manual メニューで目標状態決定
3. **冪等性チェック**: 既に期待値と一致するプロファイルは Skip（`Skipped` + `Verified=$true` で返却）
4. 変更必要なプロファイルのみ `Set-NetFirewallProfile` で適用
5. **Post-Apply Verification**: `Get-NetFirewallProfile` で再取得し期待値一致を検証
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- 管理者権限必須（`Set-NetFirewallProfile` 実行のため）
- Manual mode では複数選択（カンマ区切り）や `[4] All OFF` / `[5] All ON` の一括操作にも対応
- ルール単位の制御が必要な場合は `firewall_rule_config` / `firewall_rule_make_config` を併用
- 併用時の推奨順序: `firewall_rule_config (Import)` → `firewall_rule_make_config` → `firewall_config`（profile on/off は最終調整）

## 検証
Post-Apply Verification は **実装あり**。`Set-NetFirewallProfile` 後に `Get-NetFirewallProfile` で 3 プロファイルの現状態を読み返し、CSV／Manual で指定した期待値と完全一致するか検証します。`-Verified` 付きで `New-BatchResult` に返却。Step 3 の冪等性スキップ時は早期 return で `New-ModuleResult -Status Skipped -Verified $true` を返す（既に期待状態であることが確認されているため Verified=true 確定）設計になっており、Profile レベルの状態が確実にエビデンス化されます。


<!-- ============================================================ -->
# === modules/firewall_rule_config.md ===
<!-- ============================================================ -->

# firewall_rule_config (Standard)

**カテゴリ**: Security
**メニュー名**: Firewall Rule Export / Firewall Rule Import
**VERSION**: 0.1.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Export: ファイル整合性、Import: Name-set 包含検証）
**サブスクリプト**:
- `firewall_rule_export.ps1` … `netsh advfirewall export` でファイアウォールポリシー全体のスナップショット保存
- `firewall_rule_import.ps1` … `netsh advfirewall import` でポリシー全体を復元（**破壊的・全置換**）

## 目的
Windows ファイアウォール全体（rule + profile 状態 + logging 設定 + IPsec）を **`.wfw` 形式の真実源** として丸ごとバックアップ／復元するモジュールです。`netsh advfirewall export/import` を使い、人間可読なサイドカー（`rules.json` / `rules_show.txt` / `profiles.json` / `manifest.txt` / `rule_names.txt`）を併産することで、バイナリ `.wfw` の中身を監査可能にしています。Import は **`IAcknowledgeReplace=1` が無いと暴発しない** ゲート設計で、AutoPilot 中でも誤って全ポリシー置換が走るリスクを排除しています。Export 成功時は CSV 末尾に対応 Import 行（`Enabled=0`, `IAcknowledgeReplace=0`）を自動追記して復元動線を作る運用補助も搭載。

## 入力 (CSV)
`firewall_rule_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Mode` … `Export` / `Import`
- `SourcePath` … Import 時の `.wfw` パス（UNC 可、相対パスは `module\backup\` 基準で解決、ディレクトリ指定なら `policy.wfw` を自動付与）
- `DestinationPath` … Export 時の保存先（空欄なら `module\backup\<yyyyMMdd_HHmmss>\`）
- `IAcknowledgeReplace` … Import 時の **暴発防止承認フラグ**（0=拒否 / 1=実行許可）
- `Description` / `Segment`

## 主要ステップ
[Export]
1. CSV 読み込み（Enabled=1 + Mode=Export）
2. Pre-flight（管理者権限、`netsh` 利用可否）
3. Dry-run 表示
4. 実行確認（AutoPilot 時は自動 Y）
5. `netsh advfirewall export` 実行 + サイドカー（rules.json / profiles.json / rules_show.txt / rule_names.txt / manifest.txt with SHA256）出力
5.5 **Verification**: `policy.wfw` 存在 + サイズ > 1KB / `manifest.txt` 存在 / `rule_names.txt` 行数 == manifest 記録数
6. CSV 末尾に Import 行を自動追記（Enabled=0, IAcknowledgeReplace=0）+ `New-BatchResult -Verified`

[Import]
1. CSV 読み込み（Enabled=1 + Mode=Import + IAcknowledgeReplace=1）
2. Pre-flight + 相対パス解決
3. manifest 読み込み + OS 版差異の警告
4. 実行確認
5. `netsh advfirewall import` 実行
5.5 **Verification**: `rule_names.txt` の各 Name が Import 後 `Get-NetFirewallRule` に **包含されているか**（after が expected の superset で OK）。旧 backup なら count-only fallback
6. `New-BatchResult -Verified`

## 注意点・運用メモ
- **管理者権限必須**（両スクリプト）
- `netsh advfirewall import` は **全ポリシーを置換**（追加ではない）
- 取り込み元 `.wfw` の OS 版が現在 OS 版と異なると部分失敗の可能性あり（manifest との照合で警告）
- 相対パス先頭に `backup/` を付けると重複展開で not found（操作者向け注意点）
- Excel で CSV を開いた状態で Export を実行するとファイルロックで Import 行追記失敗 → 警告のみで Export は成功扱い

## 検証
Post-Apply Verification は **実装あり**、両モードで `-Verified` を `New-BatchResult` に返却。
- **Export**: `policy.wfw` 存在 + サイズ > 1KB、`manifest.txt` 存在、`rule_names.txt` 行数 == manifest 記録数。**現在の `Get-NetFirewallRule` count とは比較しない**（Windows 動的変動を意識した設計判断）
- **Import**: `rule_names.txt` の Name 集合が Import 後にすべて存在するか **包含検証**（後から mpssvc / AppX / GPO が動的 rule を追加しても影響なし）。1 つでも欠損すれば FAIL（欠損 Name を最大 5 個報告）
- 旧 backup（`rule_names.txt` 不在）は count-only fallback。Windows 動的変動で誤 fail 可能性あり（弱い検証）


<!-- ============================================================ -->
# === modules/firewall_rule_make_config.md ===
<!-- ============================================================ -->

# firewall_rule_make_config (Standard)

**カテゴリ**: Security
**メニュー名**: Firewall Rule Maker
**VERSION**: 0.1.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（DisplayName 取得 + 主要属性比較）
**サブスクリプト**: なし

## 目的
CSV に記述された定義から **個別の Windows ファイアウォール rule を `New-NetFirewallRule` で作成** するモジュールです。`firewall_rule_config` がポリシー全体を `.wfw` で丸ごとバックアップ／復元する **マクロ視点** を担うのに対し、本モジュールは「site-specific な rule を 1 本ずつ追加」する **マイクロ視点** を担います。冪等性は同一 `DisplayName` の rule 既存検出で SKIP、`Name` 列を明示すると内部 GUID 名で厳密マッチが効くため、後の削除や更新が確実に追跡可能になります。CSV 24 列で `Profile` / `Protocol` / `LocalPort` / `RemoteAddress` / `Program` / `Service` / `InterfaceType` / `LocalUser` (SDDL) / `EdgeTraversalPolicy` まで網羅し、`New-NetFirewallRule` の主要パラメータを CSV 1 行で表現できる広さがあります。

## 入力 (CSV)
`firewall_rule_make_list.csv` の 24 列（要約）:
- **必須**: `Enabled`, `DisplayName`, `Direction`, `Action`
- **ID**: `Name`（GUID 様。未指定は自動生成。指定時は冪等性チェックを Name で実施）
- **推奨**: `RuleEnabled`, `Profile`（複数 `;` 区切り）, `Protocol`, `LocalPort`, `RemoteAddress`
- **詳細**: `Description`, `Group`, `RemotePort`, `LocalAddress`, `IcmpType`, `EdgeTraversalPolicy`
- **対象**: `Program`, `Service`, `InterfaceType`, `InterfaceAlias`
- **セキュリティ（SDDL）**: `LocalUser`, `RemoteUser`, `RemoteMachine`
- **fabriq 標準**: `Segment`

セル内複数値は CSV 区切り `,` と衝突しないよう **`;` 区切り**（例: `Profile=Domain;Private`、`LocalPort=80;443`）。

## 主要ステップ
1. CSV 読み込み（Enabled=1）
2. 各行を validation（Direction / Action / Profile / Protocol / Port 解釈可否） + 既存 rule 重複チェック
3. plan を `[CREATE]` / `[SKIP - exists]` / `[INVALID]` で分類してプレビュー
4. 実行確認（AutoPilot 時は自動 Y）
5. `[CREATE]` のみ `New-NetFirewallRule` を splat で呼出
5.5 **Post-Apply Verification**: 作成 rule ごとに存在 / Name / Direction / Action / Enabled / Profile / EdgeTraversalPolicy を比較
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- **管理者権限必須**（`New-NetFirewallRule` のため）
- 冪等性: 同 `DisplayName` で複数 rule が存在する場合も SKIP（重複は安全側で touched しない）
- `Program` パスが存在しない場合は警告のみで rule 作成は継続（Windows は不在パスでも rule 作成を許可）
- 不正な行は実行前 reject（Fail カウント加算、他行は継続）
- `Service` 列を使う場合 Direction との組み合わせ制約あり（cmdlet が例外を返すケース。試行錯誤推奨）
- 3 ファイアウォール系モジュールの併用順序（推奨）:
  1. `firewall_rule_config (Import)` でベースライン restore
  2. `firewall_rule_make_config` で site-specific rule 追加（本モジュール）
  3. `firewall_config` で profile on/off の最終調整

## 検証
Post-Apply Verification は **実装あり**。Step 5.5 で作成された rule ごとに `Get-NetFirewallRule -Name <created.Name>` を呼び、以下を比較:
- 存在
- Name（CSV 指定時）
- Direction / Action（CSV 値と一致）
- Enabled（CSV `RuleEnabled` と一致、未指定時は True 想定）
- Profile（CSV 値を集合として比較、順序非依存）
- EdgeTraversalPolicy（CSV 指定時）

全 PASS なら `-Verified $true` で `New-BatchResult` に返却。**`InterfaceType` / `InterfaceAlias` / `LocalUser` / `RemoteUser` / `RemoteMachine` は v1 verification の対象外**（別 cmdlet `Get-NetFirewall*Filter` 経由で読まないと取れないため、cmdlet が正しく適用したことを信頼する設計）。必要なら `Get-NetFirewallInterfaceTypeFilter -AssociatedNetFirewallRule <rule>` で手動確認。


<!-- ============================================================ -->
# === modules/generic_batch_runner.md ===
<!-- ============================================================ -->

# generic_batch_runner (Standard)

**カテゴリ**: Scripts
**メニュー名**: Batch Runner
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（バッチ内部処理は観測不能）
**サブスクリプト**: `batch_runner.ps1`（メイン処理 1 本のみ。`assets/` 配下に対象 .bat を同梱）

## 目的
任意の `.bat` ファイル群を CSV 定義に従って順次実行する汎用ランナー。
事前ダウンロード済みドライバー導入バッチや顧客固有のレガシー .bat 資産を
そのまま fabriq の Profile フローに差し込みたいケース向けで、CSV にタイムアウトと
成功判定 ExitCode を持たせ、複数 .bat を一括で結果集計します。
バッチ内部のロジックには介入しないため「ブラックボックス資産の橋渡し」用途に最適です。

## 入力 (CSV)
- `Enabled`: 有効フラグ（1=実行, 0=スキップ）
- `Description`: 表示用の説明
- `BatchPath`: .bat ファイルパス（相対は `$PSScriptRoot` 基準、絶対パスも可）
- `Arguments`: バッチへ渡す引数（省略可）
- `TimeoutSec`: タイムアウト秒数（0 / 空欄 = 無制限）
- `SuccessCodes`: 成功とみなす ExitCode（カンマ区切り、省略時は 0 のみ）
- `Encoding`: エンコーディング（保留列、現状は表示のみ）
- `Segment`: Segment フィルタ（共通機構が暗黙適用）

## 主要ステップ
1. `batch_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. 実行対象一覧の存在確認とドライラン表示（[NOT FOUND] は警告）
3. 実行確認（`Confirm-ModuleExecution`、AutoPilot は自動 Y）
4. `cmd /c "path"` を `Start-Process -NoNewWindow -PassThru` で起動、PID 表示
5. `TimeoutSec` 経過時は `Stop-Process -Force` で強制終了し失敗扱い
6. ExitCode を `SuccessCodes` リストと照合し成否判定
7. `New-BatchResult` で Success/Skip/Fail を集計返却

## 注意点・運用メモ
- バッチ自身が管理者権限を必要とする場合、Profile 全体を管理者で起動する必要あり
- `Encoding` 列は予約済みだが、現実装では cscript/cmd 起動時の codepage 制御は未実装
- 多重起動防止はバッチ側責務。再実行による副作用がある .bat は冪等化を推奨
- 出力はリダイレクトせずコンソール直行のため、ログ取得はバッチ側で `>>` に書く

## 検証
Post-Apply Verification は未実装。バッチ内部の事後状態（レジストリ／ファイル／サービス）は
モジュール側からは不可視のため、`-Verified` パラメータは渡さず、実行履歴の Verified 列は空欄。
詳細検証が必要な場合は、各 .bat に `reg query` / `sc query` 等の自己検証ステップを組み込み、
ExitCode 経由で fabriq に結果を伝える設計が推奨。


<!-- ============================================================ -->
# === modules/generic_process_runner.md ===
<!-- ============================================================ -->

# generic_process_runner (Standard)

**カテゴリ**: System
**メニュー名**: Process Runner
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（起動 EXE の処理内容は観測不能）
**サブスクリプト**: `process_runner.ps1`（メイン処理 1 本のみ）

## 目的
任意の EXE を CSV 定義から順次起動する汎用ランナー。
Office 更新（OfficeC2RClient.exe）、サイレントインストーラ、コマンドラインツール等、
EXE パスと引数で完結する処理全般に使用します。
親プロセスが即終了して子プロセスが本処理を続けるブートストラップ型 EXE 向けに
`WaitProcessName` でポーリング待機する仕組みを持ち、`generic_batch_runner` の EXE 版に相当します。

## 入力 (CSV)
- `Enabled`: 有効フラグ
- `Description`: 表示用説明
- `ExecutablePath`: EXE パス（絶対 / 環境変数展開 / 相対のいずれも可）
- `Arguments`: 起動引数
- `WorkingDirectory`: 作業ディレクトリ（空欄時はモジュールフォルダ、相対 EXE パスの解決基点にも使用）
- `TimeoutSec`: タイムアウト秒数（0 / 空欄 = 無制限）
- `SuccessCodes`: 正常 ExitCode（カンマ区切り、例 `0,3010`）
- `NoNewWindow`: ウィンドウ制御（1=コンソール内実行、空欄/0=新規ウィンドウ）
- `WaitProcessName`: ポーリング待機する子プロセス名（拡張子なし）
- `Segment`: Segment フィルタ

## 主要ステップ
1. `process_list.csv` 読み込み + 環境変数展開
2. 各エントリの存在確認とドライラン表示（パス / Timeout / SuccessCodes / Window / WaitFor を表示）
3. 実行確認（AutoPilot は自動 Y）
4. `Start-Process` で EXE 起動 + PID 表示
5. `Wait-Process -Timeout` でタイムアウト監視。超過時は `Stop-Process -Force`
6. `WaitProcessName` 指定時は 5 秒間隔で当該プロセスの消滅をポーリング待機
7. ExitCode を SuccessCodes と照合 → `New-BatchResult` で集計

## 注意点・運用メモ
- パス解決順は「環境変数展開 → 絶対パスならそのまま → 相対なら WorkingDirectory or `$PSScriptRoot`」
- GUI アプリは `NoNewWindow=0`、コンソールツールは `NoNewWindow=1` 推奨
- ブートストラップ型 EXE（OfficeC2RClient 等）は必ず `WaitProcessName` を設定し、
  早すぎる「完了判定」で次モジュールに進む事故を防ぐ
- 冪等性なし。Enabled=1 のエントリは毎回 EXE が再起動されるため、
  多重実行リスクは EXE 側で処理する設計

## 検証
Post-Apply Verification は未実装。EXE 起動後の状態変化はモジュール側からは観測不能で、
ExitCode のみが成否判定根拠。`New-BatchResult` に `-Verified` は渡さない。
事後検証が必要な場合は別モジュール（例: Office 更新後に `office_license_config` の `/dstatus` 解析）で
状態を読み返す設計が推奨。


<!-- ============================================================ -->
# === modules/group_config.md ===
<!-- ============================================================ -->

# group_config (Extended)

**カテゴリ**: User Management
**メニュー名**: Local Group Member Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（実装可能なリファレンス記述あり、`Test-LocalGroupMemberExists` を Step 5.5 で再呼び出しすれば Verification 化できる構造）
**サブスクリプト**: なし

## 目的
ローカルグループ（Administrators, Remote Desktop Users 等）にドメイングループ／ドメインユーザー／ローカルユーザー／現在のログインユーザーをメンバーとして追加するモジュール。
特徴は `MemberType=CurrentUser` のサポートで、UAC 昇格時でも `Get-CimInstance Win32_ComputerSystem` の `UserName` から実際の対話セッションユーザーを正しく解決する（昇格前のユーザーが返る）。

## 入力 (CSV)
`group_list.csv`:
- **Enabled**: 有効フラグ
- **LocalGroup**: 追加先ローカルグループ名
- **MemberType**: `DomainGroup` / `DomainUser` / `LocalUser` / `CurrentUser`
- **MemberName**: 追加メンバー名（CurrentUser では空欄）
- **Domain**: ドメイン名（LocalUser/CurrentUser では空欄）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み
3. ドライラン表示（各行で `Test-LocalGroupExists` + `Test-LocalGroupMemberExists` を呼び `[Current]` / `[Change]` / `[ERROR]` マーカー）
4. `Confirm-ModuleExecution`
5. メンバー追加ループ（`Build-MemberName` で `Domain\User` 形式整形 → 冪等性チェック → `Add-LocalGroupMember`）
6. `New-BatchResult` 集計

## 注意点・運用メモ
- ドメインメンバー追加にはドメイン接続が必要
- 冪等性ロジックは MemberType により分岐: LocalUser は `COMPUTERNAME\Name` または `Name` の完全一致、それ以外は `*\Name` の末尾一致で判定
- CurrentUser 解決は WMI 失敗時、`USERDOMAIN ≠ COMPUTERNAME` で `USERDOMAIN\USERNAME`、それ以外で `USERNAME` にフォールバック

## 検証
未実装だが、Guide.txt 内に「Step 5.5 で `Test-LocalGroupMemberExists` を再呼び出しすれば実装可能」と明記されており構造的には準備されている。`-Verified` 未渡しで Verified 列は空欄。


<!-- ============================================================ -->
# === modules/history_destroyer.md ===
<!-- ============================================================ -->

# history_destroyer (Extended)

**カテゴリ**: Maintenance
**メニュー名**: History Destroyer
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（キャッシュ系は OS が即時に再生成し得るため false FAIL リスクから意図的に非対応）
**サブスクリプト**: なし（CSV ドリブンで 7 種類の Special ハンドラを内包）

## 目的
キッティング完了直前の最終クリーンアップとして、Windows の各種履歴・キャッシュ・一時ファイル・ブラウザデータ・Wi-Fi プロファイル等を一括破壊するモジュール。
CSV ドリブン設計で、`ActionType` 列により `DeletePath` / `ClearRegistry` / `Command`（Invoke-Expression）/ `Special`（専用ハンドラ）の 4 種類に分岐。Special には `clear-all-eventlogs` / `recycle-bin`（Shell32.dll P/Invoke の `SHEmptyRecycleBin`）/ `office-mru` / `edge-cleanup` / `chrome-cleanup` / `search-index` / `wifi-ssid` の 7 種類が用意される。

## 入力 (CSV)
`destroy_list.csv` — 削除ターゲット定義（Enabled, GroupName, TargetName, ActionType, TargetPath, IfNotFound, Description, Segment）。
`ssid_list.csv` — `wifi-ssid` ハンドラが参照する Wi-Fi プロファイル削除リスト（Enabled, SSID, Description, Segment）。

## 主要ステップ
1. P/Invoke (`SHEmptyRecycleBin`) を `Add-Type` で取り込み
2. **Step 0**: Explorer プロセス停止（ファイルロック解除）
3. CSV 読み込み + 環境変数展開（DeletePath 行のみ）
4. ドライラン表示（GroupName でグルーピング表示）
5. `Confirm-ModuleExecution`
6. メイン処理ループ: ActionType で switch 分岐、Special は `Invoke-DestroyHandler` でルーティング
7. **最終ステップ**: Explorer 自動再起動を最大 15 秒待機、復活しなければ `Start-Process explorer.exe` を強制実行
8. `New-BatchResult` 集計

## 注意点・運用メモ
- 13 系統のクリーンアップ: Explorer 履歴 / イベントログ / IME / Temp / クリップボード / DNS / ごみ箱 / Office MRU / Edge / Chrome / Search Index / サムネイル/アイコンキャッシュ / Wi-Fi
- 標準で `WSUS PingID/AccountDomainSid/SusClientId/SusClientIDValidation` の削除も含まれる（マスター展開後の重複検出回避）
- 管理者権限必須（イベントログ消去・Prefetch 削除・Wi-Fi プロファイル削除）
- Edge/Chrome のキャッシュ削除時は対応プロセスを `Stop-Process -Force`、その後 14 種類のターゲット（Cache, History, Cookies, Web Data 等）を削除
- 作業中は GUI が一時的に消える（Explorer 停止のため）

## 検証
キャッシュ系は OS が即時再生成するため読み返し検証で false FAIL を起こしやすく、設計上意図的に Verification 非対応。`-Verified` 未渡しで履歴 Verified 列は空欄。


<!-- ============================================================ -->
# === modules/hostname_config.md ===
<!-- ============================================================ -->

# hostname_config (Standard)

**カテゴリ**: Network
**メニュー名**: Change Hostname
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（保留中ホスト名のレジストリ読み返し）
**サブスクリプト**: `hostname_config.ps1`（単一スクリプト、ローカル CSV なし）

## 目的
ホスト名（コンピュータ名）を変更します。設定値は専用 CSV ではなく、
fabriq 起動時に `hostlist.csv` で選択したホストの `NewPCName` 列由来の
環境変数 `$env:SELECTED_NEW_PCNAME` のみから取得します。
hostlist 中心アーキテクチャを象徴するモジュールで、
Profile から呼ばれた際は無人で自動適用されます（hostlist 選択そのものが同意の代わり）。

## 入力 (CSV)
ローカル CSV なし。以下の環境変数のみ使用：
- `$env:SELECTED_NEW_PCNAME`: 新しいホスト名（hostlist.csv の NewPCName 列由来、空欄時は Skip）
- `$env:COMPUTERNAME`: 現行ホスト名（冪等性チェックに使用）

## 主要ステップ
1. `$env:SELECTED_NEW_PCNAME` 取得。未設定なら Skip
2. 現在ホスト名と目標値を表示
3. 冪等性チェック: 既に同名なら Skip
4. 実行確認（AutoPilot は自動 Y）
5. `Rename-Computer -NewName -Force` でホスト名変更
6. Step 5.5: Post-Apply Verification（後述）
7. `New-ModuleResult -Verified` で結果返却（再起動必要メッセージ付き）

## 注意点・運用メモ
- 管理者権限必須（`Rename-Computer` 自体が AdminGuard 対象）
- 反映は OS 再起動が必要。検証時点では `$env:COMPUTERNAME` は旧名のまま
- ホスト名規約違反（15 文字超、不正文字）はエラーで中断
- Domain 参加状態の場合、`Rename-Computer` には認証情報が必要となるが本モジュールは
  ワークグループ前提（domain_join モジュールとの順序依存）

## 検証
Post-Apply Verification は実装済み。`Rename-Computer` は再起動で反映されるため
現行ホスト名は変わらない代わりに、レジストリの保留値
`HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName` を読み返し、
目標値と一致するかで「適用が受理されたこと」を検証します。
結果は `New-ModuleResult -Verified $true/$false` で返却され、
実行履歴 / evidence_config の Verified 列に True/False が記録されます。


<!-- ============================================================ -->
# === modules/ipaddress_config.md ===
<!-- ============================================================ -->

# ipaddress_config (Standard)

**カテゴリ**: Network
**メニュー名**: IP Address Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（IP / Gateway / DNS の 3 軸読み返し）
**サブスクリプト**: `ipaddress_config.ps1`（単一スクリプト、ローカル CSV なし）

## 目的
Ethernet および Wi-Fi の IP アドレス・サブネット・デフォルトゲートウェイ・DNS を
hostlist.csv 由来の環境変数から一括設定します。
物理アダプタを `Get-NetAdapter -Physical` + `InterfaceDescription` の正規表現で
Ethernet / Wi-Fi に振り分け、netsh 経由で適用する Windows レガシースタックと
モダンスタック双方に整合性のある実装。確認プロンプトを持たず hostlist の値が
そのまま実行同意となる「データ駆動」モジュールの典型例です。

## 入力 (CSV)
ローカル CSV なし。hostlist.csv 由来の以下の環境変数を参照：
- `$env:SELECTED_KANRI_NO`, `SELECTED_OLD_PCNAME`, `SELECTED_NEW_PCNAME`: 表示用
- `$env:SELECTED_ETH_IP / SUBNET / GATEWAY`: Ethernet 設定
- `$env:SELECTED_WIFI_IP / SUBNET / GATEWAY`: Wi-Fi 設定
- `$env:SELECTED_DNS1 / DNS2 / DNS3 / DNS4`: 共通 DNS（Ethernet・Wi-Fi 両方に適用）

## 主要ステップ
1. 管理者権限チェック（`Test-AdminPrivilege`）
2. 環境変数読み込み + 設定内容表示
3. 物理アダプタ自動検出（Ethernet 用は Wi-Fi/Wireless/WLAN/802.11/Bluetooth 否定マッチ、
   Wi-Fi 用は同じ語の肯定マッチ。`InterfaceDescription` は OS locale 非依存で常に英語）
4. 実行確認は省略（hostlist 選択が同意代わり）
5. `netsh interface ip set address` / `set dns` / `add dns` で IP・GW・DNS 適用
6. Step 5.5: Post-Apply Verification（後述）
7. `New-ModuleResult -Verified` で集計返却

## 注意点・運用メモ
- IP 列が空欄のアダプタは設定スキップ（部分構成可能）
- Disabled アダプタは除外、Disconnected（ケーブル抜け）は対象に含む
- netsh はレガシー TCP/IP プロパティ GUI と同等のストアに書き込み、
  DHCP→Static 移行を自動処理
- サブネット表記は CIDR ではなく `255.255.255.0` 形式（内部で prefix 長変換）

## 検証
Post-Apply Verification は実装済み。各設定アダプタについて 3 軸を読み返し：
- IP + プレフィックス長: `Get-NetIPAddress` で完全一致
- デフォルトゲートウェイ: `Get-NetIPConfiguration` の `IPv4DefaultGateway.NextHop` 一致
- DNS サーバ: `Get-DnsClientServerAddress` の `ServerAddresses` を順序込み完全一致

全項目一致で `-Verified $true`、1 件でも不一致で `$false`、
hostlist で Ethernet/Wi-Fi 両方空欄の場合は `$null`（検証対象なし）を返却。


<!-- ============================================================ -->
# === modules/ipv6_config.md ===
<!-- ============================================================ -->

# ipv6_config (Extended)

**カテゴリ**: Network
**メニュー名**: IPv6 Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（`Set-NetAdapterBinding` の戻り値を信頼する設計）
**サブスクリプト**: なし

## 目的
ネットワークアダプターの IPv6 バインディング（`ms_tcpip6`）を有効化／無効化するシンプルなモジュール。
ワイルドカード対応の `AdapterPattern` で、OS 言語による「イーサネット」「Ethernet」名差を 1 CSV 内で吸収できる。

## 入力 (CSV)
`ipv6_list.csv`:
- **Enabled**: 有効フラグ
- **AdapterPattern**: アダプター名のワイルドカードパターン（例: `イーサネット*`, `Ethernet*`, `Wi-Fi*`）
- **IPv6State**: `0`=Disable / `1`=Enable（数値文字列で判定。`Enabled`/`Disabled` 文字列は不可）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み（手動で Enabled=1 をフィルタリング、無効行も表示用に保持）
3. 対象アダプター一覧表示（無効化されたエントリは `[DISABLED]` でグレー表示）
4. `Confirm-ModuleExecution`
5. `Get-NetAdapter | Where-Object Name -like $pattern` でマッチ → `Set-NetAdapterBinding -ComponentID ms_tcpip6 -Enabled $targetState`
6. `New-BatchResult` 集計

## 注意点・運用メモ
- AdapterPattern にワイルドカード使用可
- 言語環境で名前が変わるため日本語名と英語名の両方を登録しておくのが定石
- マッチするアダプターが 0 件のパターンは Skip 集計

## 検証
未実装。手動確認は `Get-NetAdapterBinding -Name <name> -ComponentID ms_tcpip6`。`-Verified` 未渡しで Verified 列は空欄。


<!-- ============================================================ -->
# === modules/local_user_config.md ===
<!-- ============================================================ -->

# local_user_config (Standard)

**カテゴリ**: User Management
**メニュー名**: Create Local Users / Delete Local Users
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: Create のみ実装あり（ユーザー存在 + グループ所属検証）／Delete はなし
**サブスクリプト**: `local_user_config.ps1`（作成）, `local_user_delete.ps1`（削除）

## 目的
ローカルユーザーアカウントの作成・削除を行います。
全 PC 共通の `local_user_list.csv` と PC 固有の `local_user_host_list.csv` を
union 適用するハイブリッド構造で、共通管理者アカウントは前者で一元管理しつつ、
特定機種だけにローカル運用ユーザーを足す等の柔軟な運用が可能です。
グループ所属はセミコロン区切り複数指定（例: `Users;Remote Desktop Users`）。

## 入力 (CSV)
**`local_user_list.csv`（全 PC 共通）**:
- `Enabled`, `UserName`, `Password`
- `PasswordNeverExpires`: 1=無期限, 0=有効期限あり
- `UserMayNotChangePassword`: 1=変更禁止, 0=変更可
- `Group`: 所属グループ（セミコロン区切り複数可）
- `Description`, `Segment`

**`local_user_host_list.csv`（任意・PC 固有）**:
- 上記列に加え `NewPCName`（hostlist の NewPCName と完全一致、SELECTED_NEW_PCNAME で突合）

## 主要ステップ（Create）
1. `local_user_list.csv` 読み込み
2. `local_user_host_list.csv` 存在時、SELECTED_NEW_PCNAME 一致行を append
3. ユーザー一覧表示（無効行は [DISABLED] と表示）
4. 実行確認（AutoPilot は自動 Y）
5. 作成ループ: `Get-LocalUser` で既存検出時 Skip、`New-LocalUser` で作成
   （`AccountNeverExpires=$true` 固定、PasswordNeverExpires / UserMayNotChangePassword は CSV 値反映）
6. グループ追加: `Add-LocalGroupMember` を Group 列各値ごとに実行
7. Step 5.5: Verification → `New-BatchResult -Verified`

## 主要ステップ（Delete）
1. CSV 読み込み（Enabled フィルタなし＝全行が削除候補）
2. host_list union
3. 一覧表示 + 確認
4. `Remove-LocalUser`、不在は Skip
5. `New-BatchResult` で集計（Verified なし）

## 注意点・運用メモ
- 管理者権限必須（`New-LocalUser` / `Remove-LocalUser`）
- パスワードは平文 CSV 保存。crypto レイヤ（`ENC:` プレフィクス + passphrase）対応は
  framework 共通機能で別途提供
- ユーザー作成失敗時もグループ追加は実行されないが、グループ追加失敗だけでは
  `successCount` には影響せず、警告のみ
- Delete スクリプトは Enabled フィルタを使わない（全行対象）。
  CSV を「削除対象リスト」として再利用する思想

## 検証
Create 側のみ Verification 実装。`Get-LocalUser` でユーザー存在を確認し、
`Get-LocalGroupMember` で `Group` 列指定の各グループに対し正しく所属しているかを
正規表現マッチで検証。1 件でも欠落があれば `-Verified $false`。
Delete 側は Verification 未実装で `-Verified` 引数を渡さない。


<!-- ============================================================ -->
# === modules/log_uploader.md ===
<!-- ============================================================ -->

# log_uploader (Extended)

**カテゴリ**: Evidence
**メニュー名**: Log Upload
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（robocopy ExitCode による成否判定はあるが厳密な読み返し検証なし）
**サブスクリプト**: なし

## 目的
プロファイル実行完了後に呼ばれるアップロードモジュール。`logs/` と `evidence/` を robocopy で UNC 共有またはローカルパスへ複製する。
モジュール内に専用 CSV を持たず、**fabriq 共通の `kernel\csv\log_destinations.csv` を参照**する点がアーキテクチャ上の特徴（同じ宛先設定を他モジュールとも共有可能）。finalize（プロファイル末尾）で自動実行される想定だが、Script Menu からの手動実行にも対応。

## 入力 (CSV)
`kernel\csv\log_destinations.csv`（モジュール外、共通配置）:
- **Path**: ローカルパス or UNC パス
- **Type**: `UNC` / `Local`
- **Enabled**: 有効フラグ
- **AuthUser**: UNC 認証ユーザー（省略可）
- **AuthPass**: UNC 認証パスワード（省略可、`ENC:` プレフィックス対応）
- **Description**: 説明（Segment 列なし、全セグメント共通）

## 主要ステップ
1. `kernel\csv\log_destinations.csv` を読み込み（Enabled=1 のみ）
2. `kernel\json\session.json` から MediaSerial を取得（フォルダ名に使用）
3. 宛先フォルダ名を生成（Unified モード: evidence ルートのフォルダ名 / Fallback: `タイムスタンプ_ホスト名_シリアル番号`）
4. `logs/` と `evidence/` の中身有無確認（両方空なら Skipped 終了）
5. トランスクリプト一時停止（ログファイルロック解除）
6. 各宛先に対して: `net use` で UNC 認証 → `New-Item` で宛先フォルダ作成 → `robocopy /E /NJH /NJS /NDL /NP /R:2 /W:1` で logs/, evidence/ をコピー → `session.json` をメタデータとしてコピー
7. `finally` で `net use /delete` 実行（必ず切断）+ `AuthPass` を null クリア
8. トランスクリプト再開
9. 全件成功なら Success / 一部失敗なら Partial / 全失敗なら Error

## 注意点・運用メモ
- finalize 時に自動呼び出しされる（Profile 末尾の標準モジュール）
- robocopy オプション `/R:2 /W:1` で再試行 2 回・待機 1 秒（一時的なネットワーク不安定に強め）
- AuthUser/AuthPass の片方だけ空欄なら認証スキップして現在のユーザーコンテキストでアクセス
- `ENC:` 暗号化パスワードは `kernel/common.ps1` の復号処理に委譲

## 検証
robocopy ExitCode 単位で成功/失敗集計し、`New-ModuleResult` で `Success` / `Partial` / `Error` を返却。`-Verified` 未渡しで Verified 列は空欄。手動検証は宛先側のフォルダ（タイムスタンプ_ホスト名_シリアル番号）を直接確認。


<!-- ============================================================ -->
# === modules/manual_kitting_assistant.md ===
<!-- ============================================================ -->

# manual_kitting_assistant (Extended)

**カテゴリ**: ManualWorks
**メニュー名**: Manual Kitting Assistant
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（手動操作の結果は OS 側の個別モジュールで検証する設計分離）
**サブスクリプト**: なし（`prompt/` フォルダに手順テキストファイル `.txt` を別途配置）

## 目的
レジストリ/CLI で自動化できない手動キッティング作業を、画面右下に常駐する WinForms モーダル風アシスタントで案内するモジュール。
1 ステップ = 1 画面（Step ID, タイトル, 手順説明, 実行ボタン, コピー1/2/3, 貼り付け, 完了）の構成で、operator はボタンクリックでアプリ起動・クリップボードコピー・貼り付けを行いながら次へ進む。
GUI ポリシーは「**マウスのみ**、F2 キーのみ完了に割当」「クリックしてもフォーカスを盗まない (`WS_EX_NOACTIVATE` + `WM_MOUSEACTIVATE` 2 段ガード)」「Ctrl+V は C# `keybd_event` ラッパー経由で送信し PowerShell マーシャリング問題を回避」。
カラーテーマは「Gundam Light」（mecha white 背景 + federation blue ヘッダー + V-fin gold アクセント + 各種ガンダムカラーボタン）。

## 入力 (CSV)
`step_list.csv`:
- **Enabled**: 有効フラグ
- **StepID**: ステップ ID（例 `S001`）
- **StepTitle**: GUI ヘッダー表示名
- **PromptFile**: `prompt/` 配下の手順説明テキストファイル名（空欄ならコンパクトモード = 説明エリアなし）
- **OpenCommand**: 実行ボタンで起動するコマンド
- **OpenArgs**: コマンド引数
- **Copy1** / **Copy2** / **Copy3**: 各コピーボタンでクリップボードに送る値
- **Segment**: Segment フィルタ

## 主要ステップ
1. `System.Windows.Forms` / `System.Drawing` ロード + C# `FabriqCtrlCSender` / `NoActivateButton` / `NoActivateForm` を `Add-Type`
2. CSV 読み込み + `prompt/` ディレクトリ存在確認 + 全 PromptFile の実在検証（不足あれば事前 Error 終了）
3. `Confirm-ModuleExecution`
4. WinForms ウィンドウを画面右下に配置（modeless, TopMost）
5. メインループ: 各ステップで `Update-StepDisplay` で UI 更新 → `DoEvents` ポーリング待機 → 完了 or キャンセル待ち
6. キャンセル時は MessageBox で確認、Yes なら以降 Skip 集計
7. 全ステップ完了後 `New-BatchResult` 集計

## 注意点・運用メモ
- 対話型セッション必須（GUI 描画必要）
- AutoPilot 運用には不向き（GUI 表示後は operator のクリック待ちでブロックされる）。AutoPilot プロファイルからは除外推奨
- PromptFile が空のステップはコンパクトモード（フォーム高さを動的縮小）で表示
- 「貼り付け」ボタンは `GetForegroundWindow` で押下時のフォアグラウンドウィンドウを記憶 → `SetForegroundWindow` でフォーカス戻し → `keybd_event` で Ctrl+V を送出する 3 段構成
- コピー成功時はボタンを 600ms グリーンフラッシュさせる視覚フィードバック
- 全ボタンに `NoActivateButton` を使い、フォーム自体も `WS_EX_NOACTIVATE` でアクティブ化を抑制（操作中の他ウィンドウのフォーカス維持）

## 検証
未実装。手動操作の結果検証は OS 側の個別モジュール（`reg_hkcu_config` 等）の責務とする設計分離。`-Verified` 未渡しで Verified 列は空欄。


<!-- ============================================================ -->
# === modules/network_profile_config.md ===
<!-- ============================================================ -->

# network_profile_config (Extended)

**カテゴリ**: Network
**メニュー名**: Network Profile Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Step 5.5 で全対象プロファイルの Category を読み返して期待値と比較、`-Verified` を返却）
**サブスクリプト**: なし

## 目的
HKLM 配下のネットワークプロファイル `NetworkList\Profiles\{GUID}` に対し、ProfileName で GUID サブキーを動的解決し `Category` 値（0=Public / 1=Private / 2=Domain）を統一的に書き込むモジュール。
マシンごとに変わる GUID を意識せず、CSV で「ProfileName + MatchMode」を指定するだけで全マッチプロファイルへ一括適用できる。
特徴的な機能は **MatchMode=All** によるベースライン + 例外運用パターン: 「まず全プロファイルを Private に」+「Guest-WiFi だけ Public に上書き」を 1 CSV で表現可能（CSV 行順で後勝ち）。

## 入力 (CSV)
`network_profile_list.csv`:
- **Enabled**: 有効フラグ
- **ProfileName**: 検索対象プロファイル名（MatchMode=All では無視）
- **Category**: `0`=Public / `1`=Private / `2`=Domain
- **MatchMode**: `Exact`（デフォルト） / `Wildcard`（PowerShell `-like` 演算子） / `All`（全プロファイル）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み + MatchMode 空欄を `Exact` に正規化
3. `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles` 存在確認 + Category/MatchMode 値検証
4. ドライラン表示（各行ごとにマッチプロファイル一覧、`[APPLY]`/`[CURRENT]`/`[NOT FOUND]` 色分け）
5. `Confirm-ModuleExecution`
6. 適用ループ（各プロファイルに対し冪等性チェック → `Set-ItemProperty Category` 書き込み）。`$finalExpected` ハッシュに後勝ちで最終期待値を記録
7. **Step 5.5 Post-Apply Verification**: `$finalExpected` の各エントリを `Test-CategoryValueMatch` で読み返し検証 → `verifyPass` / `verifyFail` 集計
8. `New-BatchResult` に `-Verified $verified` を渡して返却

## 注意点・運用メモ
- 反映には再起動不要（即時反映）
- Category=2 (Domain) は通常 NLA が自動管理する領域。手動設定後にドメイン参加状況等で Windows が再上書きする可能性あり
- 同じ ProfileName を複数行で指定した場合、CSV 行順で後勝ち（`$finalExpected` がオーバーライドルールを保持して Verification も後勝ち期待値で行う）
- 新規プロファイルの作成は行わず、既存プロファイルの Category 値変更のみ
- CSV は UTF-8 BOM 付き必須（日本語 Description / ProfileName 対応）
- SetupComplete.cmd 経由で OOBE 直後実行も可能（Guide.txt に記載例あり）

## 検証
**Post-Apply Verification 実装の好例モジュール**。Step 5.5 で書き込み後の Category を全対象プロファイルで読み返し、後勝ちルールを適用した `$finalExpected` と比較。失敗 0 件で `$verified=$true`、対象 0 件で `$null`、失敗ありで `$false` を返却。


<!-- ============================================================ -->
# === modules/odt_config.md ===
<!-- ============================================================ -->

# odt_config (Standard)

**カテゴリ**: Applications
**メニュー名**: ODT Install
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（setup.exe ExitCode のみが判定根拠）
**サブスクリプト**: `odt_install.ps1`（メイン処理 1 本。`assets/setup.exe` + 製品別 `assets/<dir>/configuration.xml` を同梱）

## 目的
Office Deployment Tool 経由で Microsoft 365 / Visio / Project 等を導入するモジュール。
エントリ単位で Offline（事前ダウンロード済み Office\）と Online（CDN ダウンロード）を
切り替えられるよう、構成 XML の `<Add SourcePath>` 属性を実行時に書き換える設計です。
既存 Click-to-Run Office を検出すると共存不可のため Error で中止。
ODT ログは `evidence/odt_log/` に hostname プレフィクスで自動収集されます。

## 入力 (CSV)
`odt_list.csv`:
- `Enabled`: 有効フラグ
- `XmlFileName`: ODT 構成 XML のファイル名（AssetsFolder 配下で解決）
- `Description`: 表示用
- `AssetsFolder`: エントリ固有 assets フォルダ（相対は `$PSScriptRoot` 基準、空欄時は `assets\`）
- `Mode`: `Offline`（既定）/ `Online`
- `Segment`

## 主要ステップ
1. `odt_list.csv` 読み込み（Enabled=1 のみ）
2. ドライラン: 各エントリの XML / AssetsFolder 存在確認、`assets\setup.exe` 必須
3. Step 3.5: 環境事前チェック + クリーンアップ
   - Office 系プロセス（WINWORD, EXCEL, OUTLOOK, ONENOTE, MSPUB, MSACCESS,
     VISIO, LYNC, Teams, OfficeClickToRun, OfficeC2RClient）強制終了
   - `ClickToRunSvc` 停止
   - ストア版 Office AppX（OneNote, Office.Desktop 等）を AppX + Provisioned 双方から削除
   - `msiserver` が Disabled なら Manual に昇格
   - System ドライブ空き容量確認（10GB 未満は Warning）
   - 既存 C2R Office 検出 → あれば Error 中止
4. 実行確認（AutoPilot は自動 Y）
5. エントリループ:
   - 構成 XML を `[xml]` で読み込み、Mode に応じて `SourcePath` 書換え/除去
   - 一時 XML を `%TEMP%` に保存し `setup.exe /configure "<temp_xml>"` を `-Wait` 同期実行
   - finally で一時 XML 削除 + ODT ログ収集（`C:\Windows\Temp\{COMPUTERNAME}-*.log` を
     セッション ts プレフィクス付きで `evidence/odt_log/` にコピー）
6. 集計返却（全成功=Success, 一部失敗=Partial, 全 XML 不在=Skipped, それ以外=Error）

## 注意点・運用メモ
- 既存 C2R Office があると Error 中止。SaRA で事前アンインストール必須
- Offline モードは AssetsFolder 配下に `Office\` オフラインソースが必須
  （本モジュールは `setup.exe /download` 機能を持たない）
- Online モードは Microsoft CDN へ HTTPS 疎通必須、数 GB 単位の通信が発生し
  setup.exe は Wait 同期のため数十分要することあり
- ストア版 Office を Cleanup フェーズで自動削除する点に注意（ユーザーデータ影響を事前確認）
- Acrobat や独自 EXE は `generic_process_runner` で別途実行する設計分離

## 検証
Post-Apply Verification は未実装。`setup.exe /configure` の ExitCode のみが成否判定で、
`New-ModuleResult` に `-Verified` は渡さない。アクティベーション状態は
別モジュール `office_license_config` の `/dstatus` 解析で別途検証する設計。
詳細トラブルシュートは `evidence/odt_log/` に自動収集された ODT ログで実施。


<!-- ============================================================ -->
# === modules/office_license_config.md ===
<!-- ============================================================ -->

# office_license_config (Standard)

**カテゴリ**: Security
**メニュー名**: Install Office Product Key / Activate Office License
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: install=なし／auth=スクリプト内で `/dstatus` 再解析（事実上の検証、`-Verified` は未渡し）
**サブスクリプト**: `office_license_install.ps1`（プロダクトキー登録）, `office_license_auth.ps1`（ライセンス認証）

## 目的
Microsoft Office のプロダクトキー登録（`OSPP.vbs /inpkey`）と
ライセンス認証（`OSPP.vbs /act`）を行う 2 スクリプト構成のモジュール。
OSPP.vbs パスは C2R 64/32bit、MSI 64/32bit の 4 候補を優先順で自動検出し、
特殊環境向けに CSV から OsppPath を強制指定可能なハイブリッド方式です。
ENC: プレフィクスによる暗号化キーの保管に対応し、複数製品（Office 本体 + Visio +
Project）を 1 度の認証で一括処理します。

## 入力 (CSV)
`office_key.csv`:
- `Enabled`: 有効フラグ
- `ProductKey`: `XXXXX-XXXXX-XXXXX-XXXXX-XXXXX` 形式（ENC: プレフィクス対応）
- `ActivationType`: `MAK` / `KMS`（疎通チェック分岐に使用）
- `OsppPath`: OSPP.vbs フルパス（空欄=自動検出）
- `Description`, `Segment`

## 主要ステップ（install）
1. 管理者権限チェック → 失敗で Error 終了
2. `office_key.csv` 読み込み（Enabled=1 のみ）
3. OSPP.vbs 自動検出。失敗時に CSV の OsppPath が皆無なら Error 終了
4. ドライラン表示（[APPLY] / [INVALID KEY] / [ENC ERROR] / [OSPP NOT FOUND] でマーキング）
5. 実行確認（AutoPilot は自動 Y）
6. 各キーごとに ENC 残存ガード → 形式バリデーション → `cscript //Nologo OSPP.vbs /inpkey:<KEY>` 実行
7. `New-BatchResult` で集計

## 主要ステップ（auth）
1. 管理者権限チェック
2. OSPP.vbs 検出
3. CSV から ActivationType を判定し MAK の場合のみ疎通チェック
   - 第一段: `Test-NetConnection activation.sls.microsoft.com:443`
   - 失敗時: `Test-Connection 8.8.8.8` フォールバック（ping 成功なら Warning で続行＝プロキシ環境想定）
   - 両方失敗で Error 中止
4. `/dstatus` で現状ステータス取得 + 全製品の LICENSE STATUS 解析
5. 全製品が LICENSED なら冪等 Skip（複数製品ロバスト判定）
6. 実行確認 → `cscript /act` 実行
7. 3 秒待機後 `/dstatus` 再実行 → 検証 → Status / Partial / Error 判定

## 注意点・運用メモ
- 管理者権限必須
- MAK: インターネット必須（プロキシ環境は ping フォールバックで継続判断）
- KMS: 社内 KMS サーバー到達性のみで OK、180 日サイクル自動再認証
- ENC: プレフィクスのまま到達した場合は passphrase 未入力検出として Error 計上
- `office_license_install.ps1` の `Find-OsppVbs` 関数は両スクリプトでローカル定義
  （`_office_common.ps1` に切り出す案は guide にコメントあるが現状未統合）

## 検証
- install: 検証は実装せず、cscript の ExitCode のみで判定。`-Verified` は未渡し
- auth: スクリプト内で `/dstatus` を再実行し全製品の LICENSE STATUS を再解析、
  Status / Message で結果を表現するが `New-ModuleResult -Verified` は未使用。
  事実上の検証はスクリプト内で完結しており、
  Verified カラム連携が必要な場合は今後 `-Verified` 引数追加で対応可能


<!-- ============================================================ -->
# === modules/office_update.md ===
<!-- ============================================================ -->

# office_update (Standard)

**カテゴリ**: Maintenance
**メニュー名**: Office Update
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: スクリプト内で VersionToReport 読み返し（事実上の検証、`-Verified` は未渡し）
**サブスクリプト**: `office_update.ps1`（メイン処理 1 本のみ）

## 目的
Click-to-Run（C2R）版 Office の更新を実行し、完了を待機する専用モジュール。
`OfficeC2RClient.exe /update` を起動し、バージョンレジストリ変化・
`Scenario` レジストリ・OfficeC2RClient プロセス存在の 3 シグナルを組み合わせた
ハイブリッド検知で完了を判定します。MSI 版（2016 以前）は対象外で、
Click-to-Run 構成が検出できない場合は Skip します。

## 入力 (CSV)
`office_update_list.csv`（SettingName/Value ペア形式・Segment 列なし）:
- `Enabled`: 行有効フラグ
- `SettingName`: 設定項目名（下記 4 種）
- `Value`: 設定値
- `Description`: 説明

採用される SettingName:
- `TimeoutMinutes`: 更新完了までの最大待機分（既定 60）
- `PollIntervalSeconds`: ポーリング間隔秒（既定 10）
- `ForceAppShutdown`: 1=Office アプリを事前強制終了
- `DisplayLevel`: 1=更新 UI 表示

## 主要ステップ
1. `office_update_list.csv` から設定読込み（SettingName/Value をハッシュ化）
2. C2R 設定レジストリ確認、未インストールなら Skip
3. `OfficeC2RClient.exe` 存在確認 + 現バージョン取得 + `Wait-NetworkReady`
4. ドライラン表示（製品 / バージョン / チャネル / 各設定値）
5. 実行確認（AutoPilot は自動 Y）
6. ForceAppShutdown=1 なら Office プロセス（WINWORD, EXCEL, POWERPNT, OUTLOOK,
   ONENOTE, MSACCESS, MSPUB, VISIO）を `Stop-Process -Force`
7. `OfficeC2RClient.exe /update user displaylevel=... forceappshutdown=...` を `Start-Process`
8. 完了待機ループ:
   - シグナル 1: `VersionToReport` レジストリ変化 → 確定的に完了
   - シグナル 2: `OfficeC2RClient` プロセス存在 → 更新処理中
   - シグナル 3: `HKLM\...\ClickToRun\Scenario` サブキー存在 → 更新操作中
   - 30 秒以上経過 + プロセス無 + Scenario 無 → 処理完了
9. `VersionToReport` 読み返し: 変化あり=Success / タイムアウト=Error / 早期完了で変化なし=Skipped

## 注意点・運用メモ
- C2R 版 Office の更新は通常再起動不要
- `OfficeC2RClient.exe` のコマンドラインスイッチは Microsoft 非公式
- ForceAppShutdown=1 では未保存ドキュメントが失われるため、
  Profile での自動実行時は事前注意喚起が必要
- 完了検知の min wait（30 秒）は更新起動直後の false negative を防ぐためのバッファ

## 検証
スクリプト内で `VersionToReport` レジストリ値の事前/事後比較を行い、
バージョン変化の有無で Status / Message を決定。`-Verified` 引数は未渡しのため
実行履歴の Verified 列は空欄になるが、事実上の検証はスクリプト内で完結。
将来 `-Verified` を追加するなら `(afterVersion -ne beforeVersion)` を直接渡せる構造。


<!-- ============================================================ -->
# === modules/partition_config.md ===
<!-- ============================================================ -->

# partition_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: Partition Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（パーティション存在 + FileSystem + サイズ ±5% 許容）
**サブスクリプト**: `partition_config.ps1`（メイン処理 1 本のみ）

## 目的
PowerShell Storage コマンドレット（`Resize-Partition` / `New-Partition` / `Format-Volume`）を
使用して既存パーティションの縮小と新規パーティション作成を行うモジュール。
複数パーティション分割（C → D → E）に対応し、CD-ROM 等が削除予定のドライブレターを
占有している場合は自動で別レターに退避（Z..Q から空きを探す）する安全機構を持ちます。
hostlist 連動ではなくモジュール CSV 駆動。

## 入力 (CSV)
`partition_list.csv`:
- `Enabled`: 有効フラグ
- `DiskNumber`: 対象ディスク番号（通常 0）
- `SourceDriveLetter`: 縮小対象（C 等）
- `SourceSizeMB`: 縮小後サイズ（MB）
- `NewDriveLetter`: 新規パーティションのドライブレター
- `NewSizeMB`: 新規サイズ（0=残り全領域、最終行のみ可）
- `FileSystem`: NTFS / ReFS
- `VolumeLabel`: ボリュームラベル
- `Description`, `Segment`

## 主要ステップ
1. CSV 読み込み（Enabled=1）
2. バリデーション:
   - `NewSizeMB=0`（残り全領域）は最大 1 行 + 必ず最終行
   - `NewDriveLetter` 重複チェック
3. ドライラン表示（[APPLY] / [SKIP] / [RELOCATE]）
4. 実行確認（AutoPilot は自動 Y）
5. **Phase A 縮小**: Source ごとに重複排除し、現サイズ ≦ 目標なら Skip、
   それ以外は `Get-PartitionSupportedSize` で SizeMin チェック後 `Resize-Partition`
6. **Phase B 新規作成**: ドライブレター衝突を検知 → CD-ROM 等は `Move-ConflictingDriveLetter` で退避、
   既存パーティションなら Skip、`New-Partition` + `Format-Volume`
7. Step 5.5: Verification（Source サイズ + 各 New パーティションの存在/FS/サイズ）
8. `New-BatchResult -Verified` で集計返却

## 注意点・運用メモ
- 管理者権限必須（Storage コマンドレット全般）
- 縮小対象ドライブが他プロセスにロックされていると失敗
- CD-ROM が D: を占有している場合の自動退避は Z, Y, X, W, V, U, T, S, R, Q から
  空きを順次選択（10 候補すべて埋まっていれば Error）
- Resize の `SizeMin` 制約は Windows ファイルシステムの未移動データ依存で、
  デフラグ未実施の場合は予想より大きい数字になることがある
- 冪等性は NewDriveLetter ベースで判定（Source の縮小は ≦ チェックで保護）

## 検証
Post-Apply Verification を実装。各 Source パーティションの実サイズが目標 ±5% 以内、
各 New パーティションの存在 / FileSystem 一致 / サイズ ±5% 以内（NewSizeMB>0 のみ）を
読み返し、全合致で `-Verified $true`、不一致 1 件以上で `$false` を返却。
ディスク操作直後の OS 内部状態安定化を考慮した tolerance 設計。


<!-- ============================================================ -->
# === modules/pianist.md ===
<!-- ============================================================ -->

# pianist (Extended)

**カテゴリ**: ManualWorks
**メニュー名**: Pianist
**VERSION**: 1.6.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Phase ごとの Manual ステータス集計から `-Verified $true/$false/$null` を返却）
**サブスクリプト**: なし。Profile 単位の構造を `profiles/<profile_name>/` 配下に持つ:
- `pianist.json` — メタデータ (label, description, target_app, default_phase)
- `procedure.csv` — Phase × Step のアクション手続マトリクス
- `values.csv` — 設定値プール（wide format、行=PC 名、列=変数）
- `shortcuts.csv` — 起動先ショートカット集（v1.0 では参照のみ）
- `instructions/<PhaseID>.txt` — Phase 手順テキスト（section marker 対応）
- `screenshots/` — author 提供の見本画像置き場（Samples タブ用）

## 目的
業務アプリの GUI 操作を伴う設定作業（ブラウザのホームページ設定、Kintone 管理画面の初期化、Office アプリの個別設定など、レジストリ/CLI で自動化できない領域）を、operator が **Profile 単位**で半自動的に進めるための GUI モジュール。`autokey_template` の正統進化版。

**位置付けの歴史**: 当初 v0.2.x は `apps/pianist/` の独立アプリだったが、履歴・スコープ問題により破棄され、2026-05-02 の決定で `modules/extended/pianist/` 配下のモジュールとして再生し fabriq 統合を完成させた経緯がある。Profile の作成・編集は Fabriq Studio に一任、本モジュールは「既存 profile を operator が Phase 単位で進めて ModuleResult を返却」する実行側に専念する。

**実行経路は 2 系統**:
1. Profile CSV 経由（推奨）: `extended/pianist/pianist.ps1, Segment=<segment_name>` を Profile に並べると、kernel が `pianist_list.csv` を Segment フィルタで絞り込み 1 profile を直接 GUI に流し込む
2. Script Menu から単発実行: ダッシュボード > Execute Module > ManualWorks > Pianist でプロファイル選択ダイアログから選ぶ

**UI ポリシー**: 全操作マウスのみ（F2 完了のみ例外）。キーボードは SendKeys でターゲットアプリへ送る専用チャネルとして温存し、Run 中の race / 迷子キー事故を防ぐ。

## 入力 (CSV)
`pianist_list.csv`（モジュール直下）:
- **Enabled**: 候補に含めるか
- **ProfileName**: `profiles/` 配下のフォルダ名（1 ProfileName = 1 phase 構造）
- **Group**: 表示用グループ名（Sample / Production など自由記述）
- **Description**: 候補一覧での説明文
- **Segment**: Profile CSV 経由の segment 名

`profiles/<name>/procedure.csv` 列: PhaseID, PhaseLabel, Color (Blue/Green/Yellow/Orange/Red/Purple/Cyan/Pink/Gray の 9 色), StepNo, Action, Value, Wait, Note。
`profiles/<name>/values.csv` は wide format（NewPCName 列 + 変数列群、`*` 行が default、PC 固有行が override、`ENC:` セル単位暗号化対応）。

## アクション 10 種
`Open` / `WaitWin` / `AppFocus` / `Type` (SendKeys) / `Key` (特殊キー) / `Wait` / `Copy` / `Paste` / `Screenshot` / `Prompt` (MessageBox 介入待ち)。

## 主要ステップ
1. `pianist_list.csv` 読み込み（Segment フィルタ適用済み）
2. 参照 profile フォルダの実在検証
3. ドライラン表示
4. `Confirm-ModuleExecution`
5. プロファイル選択（1 件なら自動、複数なら `Show-PianistProfileSelector` ダイアログ）
6. `Load-PianistProfileData` で procedure/values/shortcuts/instructions を全読み込み
7. **Step 6**: メイン GUI を絶対座標 + Anchor で構築（modeless）
8. **Step 7**: form 表示 + メインループ（DoEvents ポーリング、`UserAction` で done/cancel 判定）
9. **Step 8**: form クリーンアップ
10. **Step 9**: Phase ステータス集計 → `New-ModuleResult` を `-Verified` 付きで返却

## タブ構成（v1.5.0 以降、3 タブ）
- **[Procedure]** タブ: `instructions/<PhaseID>.txt` の `[RPA]` / `[Manual]` section を解釈して 2 段組表示。上段「RPA でやること」（Run Phase が自動実行）、下段「手動でやること」（operator が目視/クリック）
- **[Samples]** タブ (since v1.5.0): `[Samples]` section で宣言された author 提供画像のサムネイル一覧。クリックで原寸ズームのモードレスビューワを開き、Pianist 本体は引き続き操作可能。タブ見出し `Samples (N)` の N は当該 Phase の見本画像数。画像欠損時は `(missing) <filename>` placeholder
- **[Values]** タブ: 当 Phase の参照変数 + `[Show all]` トグルで全変数表示。各行 `[Copy]` ボタンで `ENC:` 復号後の値をクリップボードへ転送

## section marker（since v1.4.0、`instructions/<PhaseID>.txt` 内）
- `[RPA]` — Run Phase で自動実行される操作の説明
- `[Manual]` — operator が実施する手順
- `[Variables]` — Copy Values で表示したい変数を明示宣言（procedure.csv の `$VarName` 自動抽出と union）
- `[Samples]` — `screenshots/` 配下の見本画像ファイル名 + キャプション

後方互換: section marker が 1 つもないプレーンテキストは全文を `[Manual]` として扱う（v1.3.x 以前のサンプル無修正動作）。`[RPA]` 不在時は procedure.csv の Step 一覧を fallback 表示。

## Run-time 制御（since v1.6.0）
- **Pause**: 走っている Step / Wait が完了した時点で一時停止、再押下で再開
- **Stop**: 次の安全境界（Step 終了 / Wait 中 / WaitWin polling）で実行中断、中断後の Phase は Auto=Error 記録
- **Speed: 1.0x ⇔ 1.5x** toggle: procedure.csv の Wait・WaitWin timeout・Step 後 Wait に倍率適用（内部 settle delay 200-300ms は scale 対象外）
- `Wait-PianistResponsive`: 50ms チャンクで DoEvents をポンプしながら Sleep する責任ある待機ヘルパー（Stop/Pause を mid-wait で honor）

## 二軸ステータス
各 Phase が独立して保持する 2 軸:
- **Auto**: Run Phase の自動実行結果（NotRun / Running / Success / Partial / Error）
- **Manual**: operator 判定（Unset / OK / Warning / Error / Skip）

両方を Phase 画面下部にバッジ表示。

## 集計ロジック → ModuleResult
Manual 内訳に基づき集計:
- Error >= 1 → Status=Error, Verified=$false
- Warning >= 1 or Unset >= 1 → Status=Partial, Verified=$null
- 全て OK / Skip → Status=Success, Verified=$true
- 操作キャンセル → Status=Cancelled

return 後、kernel が Write-ExecutionHistory / Capture-ScreenEvidence / Invoke-SafeCommand を通常モジュールと同じ流儀で処理し、HTML チェックリスト・evidence summary に並ぶ。

## ベストプラクティス（Guide.txt より）
- Phase 最初の Step は `AppFocus` を必須に（Phase ボタン押下時にフォーカスが Pianist に移るため）。例外: `Open` で新規ウィンドウを起こす Phase / `WaitWin` 含む Phase
- 日本語や記号の `Type` は IME 影響を受けやすいので `Paste` を優先
- ウィンドウ待機は `WaitWin` で行い `Wait` の固定スリープに依存しない
- 機密値は Studio で `ENC:` 暗号化して values.csv に
- 重要な分岐点で `Prompt` を置き operator に目視確認させる

## 制限・既知事項
- UIA / AutomationId 操作は非対応（将来検討）
- 録画機能なし
- IF/LOOP 等の制御構文なし、線形シーケンスのみ
- Mouse 系アクションなし（座標クリックは冪等性なしのため非採用）
- 並列実行不可（1 Phase 実行中は他全ボタン disable）
- Phase 状態は session-only（Pianist 終了でリセット、永続化なし）
- Tools (ad-hoc Copy/Type, Shortcuts ランチャ) は v1.0 で非搭載
- Stop/Pause は **Step の途中**（走っている SendKeys / Open / Screenshot）では効かず、次の境界で初めて応答

## 検証
**Verification 実装モジュール**（Profile 単位の operator 判定が事実上の検証）。`pianist.ps1` 末尾で各 Phase の Manual ステータスを集計し、`Error=0 ∧ Warning=0 ∧ Unset=0` で `$verified=$true`、Error あれば `$false`、それ以外（Partial）は `$null` を `New-ModuleResult -Verified` で返却。


<!-- ============================================================ -->
# === modules/power_config.md ===
<!-- ============================================================ -->

# power_config (Standard)

**カテゴリ**: Power
**メニュー名**: Power Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Win32 API による全設定値読み返し、タイムアウトは ±1 分許容）
**サブスクリプト**: `power_config.ps1`（メイン処理 1 本のみ。P/Invoke で powrprof.dll を直接呼び出し）

## 目的
Windows の電源プラン（PowerPlan）と電源モードオーバーレイ（PowerMode、Win11 のスライダー）、
そしてディスプレイ OFF / スリープ / 休止 / HDD オフ / 電源ボタン動作 / カバー閉じ動作 /
プロセッサ最小・最大状態を一括設定する集約モジュール。
ボタン/カバー設定はコントロールパネル UI と同じ Win32 API（`PowerWriteACValueIndex` 等）を
P/Invoke で直接呼び出すため、`powercfg.exe` が機能しない HP 等 OEM PC でも正しく動作します。

## 入力 (CSV)
`power_list.csv`（プロファイル単位の行ベース、Enabled=1 を最初に発見した行を自動採用）:
- `Enabled`, `ProfileName`, `Description`
- `PowerPlan`: BALANCED / HIGH_PERFORMANCE / POWER_SAVER
- `PowerMode`: BEST_PERFORMANCE / BALANCED / BEST_EFFICIENCY（オーバーレイ、空欄/`-` で変更なし）
- `Display_TurnOff_AC` / `_Battery`: ディスプレイ OFF までの分（0=無効）
- `Sleep_After_AC` / `_Battery`, `Hibernate_After_AC` / `_Battery`
- `HardDisk_TurnOff_AC` / `_Battery`
- `PowerButton_AC` / `_Battery`, `SleepButton_AC` / `_Battery`, `LidClose_AC` / `_Battery`:
  NOTHING / SLEEP / HIBERNATE / SHUTDOWN
- `Processor_MinState_AC` / `_Battery`, `Processor_MaxState_AC` / `_Battery`: %
- `Segment`

## 主要ステップ
1. CSV 読み込み + Enabled=1 の最初の行を自動選択（無ければ対話メニュー）
2. 設定内容表示 + 確認
3. PowerPlan 切替（`powercfg /S <GUID>`）+ 冪等チェック（GETACTIVESCHEME と一致なら Skip）
4. PowerMode オーバーレイ（`PowerSetUserConfiguredACPowerMode` / `DC` API）
5. 各タイムアウト設定: `PowerReadACValueIndex` / `DC` で現在値読み返し → 一致なら Skip、
   不一致なら `powercfg /CHANGE` または `/SETACVALUEINDEX` で適用
6. ボタン/カバー設定: Win32 API `PowerWriteACValueIndex` / `DC` で直接書き込み
   （powercfg がサイレント失敗する OEM 制限を回避）
7. プロセッサ設定: 同じく Win32 API で AC/DC 別々に書き込み
8. `PowerSetActiveScheme` で確定反映
9. Step 5.5: Post-Apply Verification（後述）
10. `New-ModuleResult -Verified` で結果返却

## 注意点・運用メモ
- 管理者権限必須（`#Requires -RunAsAdministrator`）
- 冪等性チェックは `powercfg /QUERY` のテキスト解析ではなく Win32 API ReadValueIndex を使用
  → OS locale 非依存（日本語 Windows でも正しく動作）
- HP 製 PC: コントロールパネル「カバーを閉じたときの動作」表示が Win32 API 経由設定後も
  更新されない既知問題あり。実際の挙動とレジストリ値は正しく設定済み（HP 固有サービスの表示問題）
- Hibernate のタイムアウトは Windows 内部調整で ±1 分の誤差が出るため tolerance あり

## 検証
Post-Apply Verification を実装。Power Plan GUID、Power Mode（AC/DC 各々）、
全タイムアウト 8 種、ボタン/カバー 6 種、プロセッサ状態 4 種を Win32 API で
読み返し、CSV 期待値と照合。タイムアウト系は ±1 分 tolerance、その他は完全一致。
全件 PASS で `-Verified $true`、1 件でも不一致で `$false`、適用対象なしで `$null` を返却。


<!-- ============================================================ -->
# === modules/ppkg_config.md ===
<!-- ============================================================ -->

# ppkg_config (Standard)

**カテゴリ**: System
**メニュー名**: PPKG Install / PPKG Uninstall
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（cmdlet 例外有無のみで判定）
**サブスクリプト**: `ppkg_install_config.ps1`（適用）, `ppkg_uninstall_config.ps1`（削除）

## 目的
プロビジョニングパッケージ（.ppkg）の適用・削除を行う汎用モジュール。
`file/` ディレクトリに配置した PPKG ファイルを CSV 定義に従って一括処理し、
PackageName での冪等性チェックや Uninstall 時の物理ファイル削除リトライも備えます。
fabriq の標準モジュールでカバーしきれないレガシー設定（Provisioning Configuration Designer で
作成した独自設定）を一括導入する逃げ道としての位置付けです。

## 入力 (CSV)
`ppkg_list.csv`（Install / Uninstall 共通）:
- `Enabled`: 有効フラグ
- `PackageName`: パッケージ識別名（`Get-ProvisioningPackage` の PackageName と一致）
- `FileName`: `file/` 内の .ppkg ファイル名
- `Description`: 表示用
- `Segment`: Segment フィルタ

## 主要ステップ（Install）
1. CSV 読み込み + Enabled=1 抽出
2. `Install-ProvisioningPackage` cmdlet 存在確認 + `file/` ディレクトリ確認
3. ドライラン: 各エントリに [INSTALL] / [REINSTALL] / [NOT FOUND] マーカー表示
   （`Get-ProvisioningPackage` で PackageName 照合、ファイルサイズも表示）
4. 実行確認（AutoPilot は自動 Y）
5. ループ: `Install-ProvisioningPackage -QuietInstall -ForceInstall -PackagePath <ppkg>`
6. `New-BatchResult` で集計

## 主要ステップ（Uninstall）
1. CSV 読み込み + cmdlet 存在確認（`Get-ProvisioningPackage` / `Remove-ProvisioningPackage`）
2. ドライラン: [INSTALLED] / [NOT FOUND] 表示（PackageId も併記）
3. 実行確認
4. ループで再クエリ（dry-run と実行間で状態変化を考慮）→
   - Phase 1: `Remove-ProvisioningPackage -PackageId` で cmdlet 削除
   - Phase 2: `$pkg.PackagePath` の物理ファイルを 5 回まで 2 秒間隔リトライ削除
   - Phase 3: cmdlet 成功 / ファイル削除のみ成功 / 既に clean / 失敗 で結果分類
5. `New-BatchResult` で集計

## 注意点・運用メモ
- 管理者権限必須
- 適用後、設定の反映にはリブートまたは再ログオンが必要な場合あり
- `PackageName` は PPKG ビルド時に Provisioning Configuration Designer で
  指定した名前と完全一致させる必要あり（`Get-ProvisioningPackage` で確認可能）
- Install は `-ForceInstall` 付きで毎回再適用される（冪等チェックは表示のみ）
- Uninstall の物理ファイルロック対策（5 回リトライ）は Spooler 等の
  PPKG 参照プロセスが残っている状況を想定

## 検証
Post-Apply Verification は未実装。`-Verified` 引数は未渡しのため実行履歴の
Verified 列は空欄。cmdlet 例外の有無のみで成否判定するため、
PPKG が「適用されたが期待通りに動作していない」ケースは検出できない。
詳細検証が必要な場合は、PPKG が変更したレジストリ/ファイルを別モジュールで
事後確認する Profile 設計が推奨。


<!-- ============================================================ -->
# === modules/printer_delete.md ===
<!-- ============================================================ -->

# printer_delete (Standard)

**カテゴリ**: Printer
**メニュー名**: Delete Printers
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`Get-Printer` で削除済み確認）
**サブスクリプト**: `printer_delete.ps1`（メイン処理 1 本。GUI 込み Windows Forms 実装）

## 目的
不要プリンタを削除するモジュールで、3 つの動作モードが共存します。
- **KeepList mode**: hostlist の Printer1..10Name + `printer_driver_config/printer_list.csv`
  TargetHost マッチ行を「残すリスト」として、それ以外の TCP/IP プリンタを削除候補化
- **Explicit mode**: `printer_delete.csv` 列挙のプリンタを名前完全一致で削除（仮想プリンタも削除可）
- **Manual mode**: フォールバック。GUI で手動選択して削除

マスタイメージに全部署プリンタを一括登録 → 各 PC は hostlist で「使うプリンタ」を宣言、
KeepList mode が自動的に他部署プリンタを掃除する部署展開シナリオが代表的ユースケース。

## 入力 (CSV)
`printer_delete.csv`（任意・Explicit mode 用）:
- `Enabled`: 有効フラグ
- `TargetHost`: 対象ホスト名（空=全ホスト、値=NewPCName 完全一致、大小文字非区別）
- `PrinterName`: プリンタ名完全一致
- `Description`, `Segment`

クロス参照: `..\printer_driver_config\printer_list.csv`（Keep List 拡張ソース）

## 主要ステップ
1. Keep List 収集
   - (a) hostlist 環境変数 `SELECTED_PRINTER_1..10_NAME`
   - (b) `printer_driver_config/printer_list.csv` の TargetHost マッチ行
2. Explicit 削除リスト収集（`printer_delete.csv` の TargetHost マッチ行）
3. インストール済みプリンタ列挙 + TCP/IP ポート（`PrinterHostAddress` 持ち）判定
4. 削除候補計算: 各プリンタに Reason ラベル（Keep / KeepList / Explicit / KeepList+Explicit / Manual）と
   PreCheck 状態を付与
5. AutoPilot 分岐:
   - AutoPilot: GUI スキップ、PreCheck=$true を一括 `Remove-Printer`、
     候補 0 件なら Skipped（全削除事故を構造的に防止）
   - Manual: ダーク基調の Windows Forms GUI（DataGridView + Select All ボタン）を表示し、
     ユーザーがチェックボックス選択 → MessageBox 確認 → 削除実行
6. Step 5.5: 削除済みプリンタを `Get-Printer -Name` で再検索し null 確認
7. `New-ModuleResult -Verified` で集計返却

## 注意点・運用メモ
- 管理者権限必須
- KeepList mode は TCP/IP ポート紐付きプリンタのみ対象。仮想プリンタ
  （Microsoft Print to PDF / OneNote / XPS / Fax）は安全装置で対象外、
  削除したい場合は Explicit mode で名前明記
- `printer_driver_config` 配下の `printer_list.csv` をクロス参照することで、
  Register Printers で登録したプリンタが自動的に KeepList にも入る二重管理防止設計
- AutoPilot では `$global:AutoPilotMode` が kernel 側で設定される
- GUI 配色は fabriq 標準のダークテーマ（bgDark=ARGB(30,30,30) 等）

## 検証
Post-Apply Verification を実装。削除実行した各プリンタ名について
`Get-Printer -Name <name>` を実行し、`$null` 返却（つまり削除済み）を確認。
1 件でも残存していれば `-Verified $false`。削除実行 0 件（GUI でキャンセル等）の場合は
Cancelled を返し、Verification はスキップ。


<!-- ============================================================ -->
# === modules/printer_driver_config.md ===
<!-- ============================================================ -->

# printer_driver_config (Standard)

**カテゴリ**: Printer
**メニュー名**: Printer Drivers / Register Printers / Uninstall Printer Drivers
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: install=実装あり／register=実装あり（4 項目突合）／uninstall=なし
**サブスクリプト**: `printer_driver_install.ps1`（ドライバ登録）, `printer_config.ps1`（プリンタ登録）, `printer_driver_uninstall.ps1`（ドライバ削除）。`tools/7z.exe` を同梱

## 目的
プリンタドライバとプリンタの登録/削除を行う 3 スクリプト構成のモジュール。
`INF/` 配下に置いた EXE/ZIP 自己解凍アーカイブを 7z.exe で自動展開し、
INF 内のドライバ名と hostlist + `printer_driver_list.csv` の宣言を完全一致照合して
自動インストールする「Auto mode」と、対話選択する「Interactive mode」を持ちます。
プリンタ登録は hostlist 環境変数（最大 10 台）と `printer_list.csv` の union 適用で
10 台超や共通プリンタ宣言にも対応します。

## 入力 (CSV)
**`printer_driver_list.csv`（任意・ドライバ拡張）**:
- `Enabled`, `TargetHost`（空=全ホスト、値=NewPCName 一致）, `DriverName`（INF モデル名と完全一致）, `Description`

**`printer_list.csv`（任意・プリンタ拡張、`printer_delete` がクロス参照）**:
- `Enabled`, `TargetHost`, `PrinterName`, `DriverName`, `PortAddress`, `Description`

hostlist 環境変数: `SELECTED_PRINTER_1..10_NAME / DRIVER / PORT`, `SELECTED_NEW_PCNAME`

## 主要ステップ（Printer Drivers / install）
0. `INF/` 直下の .exe / .zip を `tools/7z.exe` で同名フォルダに自動展開（冪等）
   - 7z 不在時は .zip のみ `Expand-Archive` でフォールバック
1. hostlist + `printer_driver_list.csv` から要求ドライバ名を union 収集
2. `INF/` 配下の全 INF をスキャンし `[Manufacturer]` + アーキテクチャモデルセクションから
   ドライバ名を抽出 → driverMap 構築
3. 要求名 vs INF 内ドライバ名を完全一致照合（[マッチ] / [Unmatched]）
4. 実行確認（Auto mode の AutoPilot は自動 Y）
5. `pnputil /add-driver <inf> /install` で Driver Store 登録 → DriverStore の oem*.inf 経路を解決
6. `Add-PrinterDriver -Name <model> -InfPath <store>` で登録
7. Step 5.5: `Get-PrinterDriver -Name` で各ドライバ存在確認

## 主要ステップ（Register Printers）
1. hostlist 環境変数 + `printer_list.csv` から printers 配列を union
2. ドライバ存在確認（`Get-PrinterDriver`）→ 不足は警告
3. 確認 → `Add-PrinterPort -PrinterHostAddress <IP>` でポート IP_<addr> 作成
4. `Add-Printer -Name -DriverName -PortName` で登録（既存はスキップ）
5. Step 5.5: `Get-Printer` / `Get-PrinterPort` で 4 項目（存在 / DriverName / PortName /
   PrinterHostAddress）を完全一致検証

## 主要ステップ（Uninstall Drivers）
1. `Get-PrinterDriver` 列挙 → Microsoft 標準を除外したリスト表示
2. 番号入力で対象選択
3. 該当ドライバを使うプリンタを事前削除
4. `Restart-Service spooler` でロック解放
5. `Remove-PrinterDriver` 実行 → `pnputil /enum-drivers` から oemN.inf を解決し
   `pnputil /delete-driver oemN.inf /force` で Driver Store からも削除

## 注意点・運用メモ
- 管理者権限必須（pnputil / Add-PrinterDriver / Add-Printer / Spooler restart 全て）
- `INF/` 直下の EXE/ZIP は冪等展開。再展開は手動でフォルダ削除が必要
- `tools/7z.exe` + `7z.dll` は GNU LGPL v2.1+ で同梱、`README-license.txt` で
  バージョン明示。`THIRD_PARTY_NOTICES.md` も連携
- INF 解析はアーキテクチャ（NTamd64 / NTx86）を `[Environment]::Is64BitOperatingSystem` で判定
- 仮想プリンタ（Microsoft Print to PDF 等）は Microsoft 標準パターンで Uninstall 対象から除外

## 検証
- install: `Get-PrinterDriver -Name <model>` で各 matched driver の Driver Store 登録を検証、
  全件存在で `-Verified $true`
- register: 各プリンタに対し Printer 名 / DriverName / PortName /
  PrinterHostAddress の 4 項目完全一致を検証、1 件でも不一致で `$false`
  （reason: not found / driver mismatch / port mismatch / binding mismatch を個別表示）
- uninstall: Verification 未実装、結果サマリのみ


<!-- ============================================================ -->
# === modules/process_killer.md ===
<!-- ============================================================ -->

# process_killer (Standard)

**カテゴリ**: System
**メニュー名**: Process Killer
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（再起動型常駐プロセス想定で読み返し検証は意図的に非実装）
**サブスクリプト**: `process_killer.ps1`（メイン処理 1 本のみ）

## 目的
指定したプロセスを強制終了するシンプルな構造のモジュール。
キッティング作業中に開いたままになりがちな Edge / Notepad / OEM 常駐ツール等を、
Profile の終盤で一括クローズしクリーンな状態にする用途。
`Get-Process -Name <ProcessName>`（拡張子なし）を入力契約とし、
既に停止している場合は Skip で冪等性を保ちます。

## 入力 (CSV)
`process_list.csv`:
- `Enabled`: 有効フラグ
- `ProcessName`: プロセス名（.exe なし。例: msedge, notepad）
- `Description`: 表示用説明
- `Segment`: Segment フィルタ

## 主要ステップ
1. `process_list.csv` 読み込み（Enabled=1 のみ、`Description` 必須）
2. ドライラン: 各エントリで `Get-Process` を実行し、[Running] / [Not Running] と
   インスタンス数を表示
3. 実行確認（AutoPilot は自動 Y）
4. 適用ループ:
   - 冪等性チェック: 該当プロセス 0 件なら Skip（再実行で副作用なし）
   - 実行中なら `Stop-Process -Name -Force` で全インスタンス強制終了
   - 例外時は失敗計上
5. `New-BatchResult` で Success / Skip / Fail 集計

## 注意点・運用メモ
- 管理者権限推奨（他ユーザー所有プロセスを終了する場合に必要）
- ProcessName は `.exe` を含めない（`Get-Process -Name` の契約）
- 常駐監視型プロセス（OneDrive, Defender 関連等）は終了直後に再起動する設計のため、
  本モジュールでは検証しない方針
- 重要なシステムプロセス（lsass, csrss 等）を指定すると OS が不安定化するリスクあり、
  CSV のレビューは慎重に

## 検証
Post-Apply Verification は意図的に未実装。
終了後すぐに同名プロセスが再起動するケース（常駐監視型サービス、autorun）を
想定した設計判断。`-Verified` を渡さないため実行履歴の Verified 列は空欄。
「kill した瞬間の事実」のみが本モジュールの責任範囲、という割り切り。


<!-- ============================================================ -->
# === modules/profile_delete.md ===
<!-- ============================================================ -->

# profile_delete (Standard)

**カテゴリ**: User Management
**メニュー名**: Delete User Profiles
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（実装推奨メモは Guide にあり、現状未実装）
**サブスクリプト**: `profile_delete.ps1`（メイン処理 1 本のみ）

## 目的
指定ユーザーのプロファイル（`%SystemDrive%\Users\<UserName>`）を削除するモジュール。
WMI（`Win32_UserProfile`）を使用して Windows のプロファイル登録を正しく解除し、
WMI レコードが存在しない孤児フォルダや WMI 削除後にフォルダだけが残ったケースに
備えて `Remove-Item -Recurse -Force` の物理削除フォールバックを持ちます。
キッティング後に testuser や defaultuser0 を一掃する用途が代表的。

## 入力 (CSV)
`profile_list.csv`:
- `Enabled`: 有効フラグ（1=削除対象, 0=スキップ）
- `UserName`: ユーザー名（`C:\Users\<UserName>` のフォルダ名と一致）
- `Description`: 表示用
- `Segment`: Segment フィルタ

## 主要ステップ
1. 管理者権限チェック（`Test-AdminPrivilege`）
2. `profile_list.csv` 読み込み（Enabled=1 のみ、UserName 必須）
3. `%SystemDrive%\Users` 存在確認（D: ドライブインストール等にも追従）
4. ドライラン: 各エントリの実フォルダ存在を [APPLY] / [SKIP] と表示
5. 実行確認（AutoPilot は自動 Y）
6. 削除ループ:
   - フォルダ不在は Skip
   - **Stage 1**: `Get-CimInstance Win32_UserProfile | Where { LocalPath -like "*\$userName" }` で
     WMI レコードを取得し `Remove-CimInstance` で正規登録解除
   - **Stage 2**: フォルダがまだ残っていれば `Remove-Item -Path -Recurse -Force` で物理削除
7. `New-BatchResult` で集計

## 注意点・運用メモ
- 管理者権限必須（`Win32_UserProfile.Remove-CimInstance` 実行）
- 現在ログオン中のユーザープロファイルは削除不可（OS が拒否、Error 計上）
- WMI 削除のみではフォルダが残るケースがあり、Stage 2 のフォールバックは必須
- 削除されたユーザーの SID / プロファイル情報が Windows から登録解除される（レジストリも含む）
- パス検証（destructive_path_safety）: `Join-Path $usersBase $userName` の結果が
  `$usersBase` 配下であることを暗黙に前提（CSV の UserName に `..\` 等の
  path traversal が混入した場合の防御は弱い、運用 CSV のレビュー前提）

## 検証
Post-Apply Verification は未実装。Guide には実装推奨メモがあり、
- `Test-Path "$env:SystemDrive\Users\<UserName>"` が `$false`
- `Get-CimInstance Win32_UserProfile` で該当 LocalPath 不在

の 2 条件で `-Verified $true` 返却が可能との設計指針が記載されているが、現状の
`New-BatchResult` には `-Verified` を渡していない。実行履歴の Verified 列は空欄。


<!-- ============================================================ -->
# === modules/reg_hkcu_config.md ===
<!-- ============================================================ -->

# reg_hkcu_config (Standard)

**カテゴリ**: Registry
**メニュー名**: Registry Config (HKCU) / Registry Delete (HKCU)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Config 側のみ。Delete 側は未実装）
**サブスクリプト**:
- `reg_hkcu_config.ps1` … HKCU + Default プロファイルへの値書き込み（メイン）
- `reg_hkcu_delete.ps1` … HKCU + Default プロファイルからの値削除
- 配置物: `C:\ProgramData\fabriq\apply_hkcu.ps1`（Active Setup / Startup Batch 有効時）

## 目的
HKEY_CURRENT_USER 配下のレジストリ値を、ログオン中ユーザーの HKCU と
`C:\Users\Default\ntuser.dat`（新規ユーザーひな型）の両方に同時適用するモジュール。
Default ハイブを load して書き込むことで、本モジュール実行後に作成される新規ユーザーは
最初から同じ既定値を持つ。Explorer 初期化による上書きが問題となるキー向けに、Active Setup
登録と Startup Batch 配置（いずれもデフォルト無効）の二重補完機構も用意されている。

## 入力 (CSV)
`reg_hkcu_list*.csv` のパターンに一致する全ファイルが自動で読み込まれ集約される
（ジャンル分割管理可能、例: `reg_hkcu_list_ui.csv`）。
- `Enabled`: 有効フラグ（1=実行 / 0=スキップ）
- `AdminID`: 管理番号（表示用）
- `SettingTitle`: 設定タイトル（表示用）
- `KeyPath`: `HKEY_CURRENT_USER\...` のフルパス
- `KeyName`: 値名（`@` で既定値）
- `Type`: REG_DWORD / REG_SZ / REG_EXPAND_SZ / REG_QWORD / REG_BINARY / REG_MULTI_SZ
- `Value`: 設定値
- `Segment`: Segment フィルタ（任意）

## 主要ステップ
1. `Resolve-HkcuRoot` でログオンユーザーの HKCU 書き込み先（`HKCU:` または `HKU:\<SID>`）を解決
2. `reg_hkcu_list*.csv` を全件読み込み → Enabled=1 を抽出
3. Default ハイブ load（`reg load HKEY_USERS\Hive C:\Users\Default\ntuser.dat`）
4. ドライラン一覧表示（各エントリを `Test-RegistryValueMatch` で `[Current]`/`[Change]` 色分け）
5. 実行確認（AutoPilot 時は自動 Y）
6. 書き込みループ（HKCU 側 → HIVE 側、`$FORCE_OVERWRITE=$false` 時は一致エントリを Skip）
7. Step 5.5: Post-Apply Verification（HKCU と HIVE 双方を再読み出し、全件一致時のみ `-Verified=$true`）→ HKU:\Hive unload（失敗時 1 回リトライ）→ 必要なら Active Setup / Startup Batch 配置 → `New-BatchResult` 返却

## 注意点・運用メモ
- 管理者権限必須（Default ハイブ load/unload、HKLM への Active Setup 登録、ProgramData 配下への配置）
- 昇格セッション対策の HKCU リダイレクトを内蔵。実際の書き込み先は `Current User` または
  `<username> (via HKU)` と表示されるため意図したユーザーへの書き込みかを確認可能
- スクリプト先頭フラグ:
  - `$FORCE_OVERWRITE`（既定 `$true`）= 常時上書き、`$false` で冪等 Skip
  - `$ENABLE_ACTIVE_SETUP`（既定 `$false`）= 新規ユーザー初回ログオン時の HKCU 補完登録
  - `$ENABLE_STARTUP_BATCH`（既定 `$false`）= Explorer 起動後の Startup ショートカット経由補完
- Active Setup / Startup Batch を有効化した場合、CSV 内容から `apply_hkcu.ps1` が自動生成されるため、
  CSV を更新したらモジュール再実行で再生成する必要がある
- ハイブ unload は `[gc]::Collect()` + 1 回リトライで failsafe。失敗時は Warning のみで処理継続
- Delete 側はスクリプト構造は Config と同じだが Verification 未実装（Verified 列は空欄）

## 検証
Step 5.5 で HKCU 側と HIVE 側の各エントリを `Test-RegistryValueMatch` で読み返し、
DWord/QWord は数値比較、Binary は HEX 比較、MultiString は改行 join 比較、その他は文字列比較で
期待値一致を判定。HIVE 側は load に成功している場合のみ検証対象。1 件でも失敗すれば
`-Verified=$false` で `New-BatchResult` に渡す。Active Setup / Startup Batch の効果（新規ユーザー
初回ログオン時の挙動）はその場では検証不可。


<!-- ============================================================ -->
# === modules/reg_hklm_config.md ===
<!-- ============================================================ -->

# reg_hklm_config (Standard)

**カテゴリ**: Registry
**メニュー名**: Registry Config (HKLM) / Registry Delete (HKLM)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Config 側のみ。Delete 側は未実装）
**サブスクリプト**:
- `reg_hklm_config.ps1` … HKLM への値書き込み（メイン）
- `reg_hklm_delete.ps1` … HKLM の値削除

## 目的
HKEY_LOCAL_MACHINE 配下のレジストリ値を CSV ベースで一括適用・削除するモジュール。
ファイアウォール無効化、Ctrl+Alt+Del 必須化解除、高速スタートアップ無効化など、
マシン全体に効くポリシー系・OS 内部設定系のレジストリを一元管理する。
冪等性ヘルパー `Test-RegistryValueMatch` を内蔵し、6 つの値種（REG_DWORD/SZ/QWORD/BINARY/
MULTI_SZ/EXPAND_SZ）ごとに型に応じた比較を行う。

## 入力 (CSV)
`reg_hklm_list*.csv` のパターンに一致する全ファイルが自動で読み込まれ集約される
（ジャンル分割管理可能、例: `reg_hklm_list_security.csv`, `reg_hklm_list_ui.csv`）。
- `Enabled`: 有効フラグ（1=実行 / 0=スキップ）
- `AdminID`: 管理番号（表示用）
- `SettingTitle`: 設定タイトル（表示用）
- `KeyPath`: `HKEY_LOCAL_MACHINE\...` のフルパス
- `KeyName`: 値名
- `Type`: REG_DWORD / REG_SZ / REG_EXPAND_SZ / REG_QWORD / REG_BINARY / REG_MULTI_SZ
- `Value`: 設定値
- `Segment`: Segment フィルタ（任意）

## 主要ステップ
1. `reg_hklm_list*.csv` を全件読み込み（複数ファイル集約）
2. Enabled=1 のみ抽出 → 0 件なら Skipped
3. ドライラン表示（`Test-RegistryValueMatch` で `[Current]`/`[Change]` 色分け、
   `$FORCE_OVERWRITE=$true` の場合は FORCE MODE 表示）
4. 実行確認（AutoPilot 時は自動 Y）
5. 書き込みループ（型別キャスト後 `Set-ItemProperty` / `New-ItemProperty`、
   `$FORCE_OVERWRITE=$false` 時は一致エントリを Skip）
6. Step 5.5: Post-Apply Verification（全エントリを再読み出しして期待値一致を確認）
7. `New-BatchResult -Verified $verified` で結果返却

## 注意点・運用メモ
- 管理者権限必須
- CSV は Shift-JIS（ANSI）保存推奨。BOM 付き UTF-8 はヘッダー読み込みで不具合の可能性
- 先頭フラグ `$FORCE_OVERWRITE`（既定 `$true`）で冪等性挙動を切り替え。`$false` 時は
  目標値と一致するエントリは `[Skip]` 扱い、Verification は常に実行
- 新規キーの作成（`New-Item -Path -Force`）と既存値の上書き両方をサポート
- Binary 値は CSV 上のアスキー HEX を `[byte[]]` に変換
- DWord/QWord は `[int]` キャスト前提（10 進数表記）
- Delete 側 (`reg_hklm_delete.ps1`) は Verification 未実装で履歴の Verified 列は空欄

## 検証
Step 5.5 で全エントリに対し `Test-RegistryValueMatch` を再実行。型別比較ロジックは
HKCU 版と同一で、DWord/QWord は数値比較、Binary は HEX 比較、MultiString は改行 join 比較、
その他は文字列比較。1 件でも失敗があれば `-Verified=$false` で `New-BatchResult` に渡す。
ファイアウォール無効化のように OS 側で同等の項目が複数経路で制御される設定でも、
このモジュールはレジストリ値の書き込み一致のみを保証し、ランタイム挙動の検証は行わない
（その層は対応する `firewall_config` 等に分担させる設計）。


<!-- ============================================================ -->
# === modules/reg_template.md ===
<!-- ============================================================ -->

# reg_template (Extended)

**カテゴリ**: System
**メニュー名**: Registry Backup (Template) / Registry Import (Template)（2 メニュー）
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（reg.exe ExitCode 判定のみ、`.reg` 内容や復元後の値読み返しなし）
**サブスクリプト**: `reg_backup.ps1`（エクスポート）+ `reg_import.ps1`（インポート）。`backup/` ディレクトリにタイムスタンプ付き `.reg` を蓄積

## 目的
任意のレジストリキーを `.reg` ファイルにバックアップ／復元するための**テンプレートモジュール**（モジュール名に "Template" を冠する通り）。`desktop_icon_config` は本モジュールを特定キー専用に固定したサンプルと位置付けられ、本テンプレートは「ディレクトリをコピー → reg_list.csv を編集して任意キー対応モジュールを量産」する出発点。

## 入力 (CSV)
`reg_list.csv`:
- **Enabled**: 有効フラグ
- **RegistryPath**: フルパス（`HKEY_LOCAL_MACHINE\...` 形式）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ（Backup）
1. CSV 読み込み
2. `backup/` 必要に応じて作成
3. 各レジストリパスの存在確認 + リスト表示（`[OK]`/`[--]` マーカー）
4. `Confirm-ModuleExecution`
5. 各キーに対し `reg.exe export` を実行、ファイル名は `yyyyMMdd_HHmmss_<sanitized_path>.reg`（パス内の `\/:*?"<>|` を `_` に置換、長すぎる場合は 80 文字に切り詰め）
6. `New-BatchResult` 集計

## 主要ステップ（Import）
1. CSV 読み込み
2. `backup/` 存在確認
3. 各 RegistryPath に対し `*_<sanitized_path>.reg` パターンでマッチする最新ファイルを `Sort-Object LastWriteTime -Descending` で特定
4. 表示（`[Ready]`/`[Missing]`、複数バックアップあれば「N backups available, using latest」表示）
5. 全件 Missing なら Error 終了
6. Warning 表示後 `Confirm-ModuleExecution`
7. 各 `.reg` を `reg.exe import`
8. `New-BatchResult` 集計（成功 1 件以上で「サインアウト/再起動が必要な場合あり」警告）

## 注意点・運用メモ
- 同一キーの過去バックアップが `backup/` 内に蓄積される（タイムスタンプで自然に世代管理）
- インポートは常に最新ファイルを使用、過去版に戻したい場合は手動でファイル名指定が必要（GUI 未提供）
- 管理者権限必須（HKLM 操作のため）
- ファイル名サニタイズの 80 文字制限により、極端に長いキーパスは衝突リスクあり

## 検証
未実装。両スクリプトとも `reg.exe` の ExitCode のみで成否判定。`-Verified` 未渡しで Verified 列は空欄。


<!-- ============================================================ -->
# === modules/resolution_api_config.md ===
<!-- ============================================================ -->

# resolution_api_config (Standard)

**カテゴリ**: Display
**メニュー名**: Resolution Config (Live)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（事前判定のみ。適用後の読み返し検証なし → `-Verified` 未渡し）
**サブスクリプト**: なし（C# 型 `ResolutionHandler` を `Add-Type` で同居）

## 目的
ディスプレイ解像度を Win32 API（`ChangeDisplaySettings` / `EnumDisplaySettings`）経由で
即時変更するモジュール。レジストリ書き込みと違って再起動なしに反映される（API が
`DISP_CHANGE_RESTART` を返した場合のみ警告）。CSV に複数候補を書いておけるが、通常運用は
Enabled=1 を 1 行にしておく形（解像度の連続切替は副作用が大きいため）。

## 入力 (CSV)
`resolution_list.csv`
- `Enabled`: 有効フラグ（1=実行 / 0=スキップ）
- `Width`: 横幅（ピクセル、整数）
- `Height`: 高さ（ピクセル、整数）
- `Description`: 説明（表示用、例 "Full HD" / "QHD"）
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 1920x1080（Full HD, Enabled=1）, 2560x1440, 1680x1050, 1600x900,
1366x768, 1280x1024, 1280x720（いずれも Enabled=0）。

## 主要ステップ
1. `Add-Type` で C# `ResolutionHandler` 構造体（DEVMODE / 各種定数）をロード
2. `resolution_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
3. `EnumDisplaySettings(ENUM_CURRENT_SETTINGS)` で現在解像度を取得・表示
4. ターゲット一覧をドライラン表示（現在値と一致するエントリは `[SKIP]`）
5. 全件 already-set なら Skipped で early return / 差分があれば実行確認（AutoPilot 自動 Y）
6. 各エントリに `ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY)` を発行
7. 戻り値で分岐: `DISP_CHANGE_SUCCESSFUL` → 成功カウント＋以後の比較用に `currentW/H` を
   更新 / `DISP_CHANGE_RESTART` → 「再起動後に反映」警告して成功扱い / `DISP_CHANGE_FAILED`
   → 失敗 → `New-BatchResult`（`-Verified` は渡さない）

## 注意点・運用メモ
- ハードウェア・ドライバ非対応の解像度は `DISP_CHANGE_FAILED` で失敗扱い
- マルチモニター環境では現状第一プライマリディスプレイのみ対象（device 名 null 指定）
- リフレッシュレートは渡していないため OS 既定が採用される
- 連続変更は前段の成功値を `currentW/H` に反映するため、CSV に複数 Enabled=1 を並べると
  「次に Enabled=1 のものに切り替わる」だけの動作になる（最終的に CSV 末尾に近い解像度が残る）
- 履歴の Verified 列は空欄になる（後述の理由）

## 検証
事前判定としては `GetCurrentResolution()` で現在値と目標値が一致する場合に `[SKIP]` を出す
冪等性チェックがあるが、`ChangeDisplaySettings` 後にもう一度読み返して
期待値一致を確認する Step 5.5 は実装されていない。
これは `DISP_CHANGE_RESTART` のように「呼び出しは成功したが反映は再起動後」の状態を
読み返しでは区別できないこと、および即時反映を保証しても次に新しい接続イベントで
解像度がリセットされうるディスプレイ実装が存在することによる設計判断。
このため `New-BatchResult` 呼び出しに `-Verified` は渡されず、実行履歴の Verified 列は空欄。


<!-- ============================================================ -->
# === modules/restart_config.md ===
<!-- ============================================================ -->

# restart_config (Standard)

**カテゴリ**: System
**メニュー名**: Restart with AutoRun
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（再起動で自プロセスが消えるため検証不可、`-Verified` 未渡し）
**サブスクリプト**: なし

## 目的
Fabriq 自身を `HKLM\...\RunOnce` に登録した上で、PC を再起動するモジュール。
Profile の最終ステップに置くことで「再起動 → 再起動後に Fabriq が自動再開」のループを成立
させる、フレームワークの中核フロー部品。本モジュールの主機能は RunOnce 登録であり、
再起動はその後の付随アクションとして連結されている。

## 入力 (CSV)
設定 CSV なし。Order/Enabled は `module.csv` のみで管理（Order=99 で Profile 末尾向け）。

## 主要ステップ
1. `restart_config` から 3 階層上 (`..\..\..`) を fabriq ルートとして解決し、`Fabriq.exe` 存在確認
2. `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce` の現在値を確認・表示
3. RunOnce 登録予定内容（Path / Name / Value）をドライラン表示
4. 実行確認（AutoPilot 時でも確認スキップしない設計 = 不可逆操作の安全策）
5. `RunOnce\FabriqAutoStart` に `"...\Fabriq.exe"` を書き込み
6. `New-ModuleResult -Status Success` を先に確定（再起動で履歴保存が間に合うように）
7. `Invoke-CountdownRestart -Seconds 10` で 10 秒カウントダウン後 OS 再起動

## 注意点・運用メモ
- 管理者権限必須（HKLM RunOnce 書き込み）
- AutoPilot 非対応（メモ `feedback_autopilot_wording.md` の通り、AutoPilot は「無人」ではなく
  オペレーター立ち会い前提のスキップだが、再起動はそれでも明示確認を要求するクリティカル例外）
- RunOnce は OS 再起動後の自動起動が成功すれば値が消えるため、リブート後の Fabriq 自動起動の
  有無で登録成否を判定する運用
- `Invoke-CountdownRestart` は kernel/common.ps1 提供。Ctrl+C キャンセル余地あり
- Order=99 だが、Profile 設計次第ではモジュール途中段階のリブートにも使える
  （Profile 上で複数回呼ぶケース）

## 検証
本モジュールに Post-Apply Verification は実装されていない。RunOnce 登録は理屈上は
レジストリ読み返しで検証可能だが、登録直後にプロセス自身が消える（OS 再起動）ため、
モジュールから戻り値を返したあとの Verified 評価フェーズが存在しないという技術的制約により
意図的に未実装。`New-ModuleResult` 呼び出しに `-Verified` を渡していないため、
実行履歴の Verified 列は空欄になる。再起動後の Fabriq 自動起動が事実上の検証手段。


<!-- ============================================================ -->
# === modules/restore_point.md ===
<!-- ============================================================ -->

# restore_point (Standard)

**カテゴリ**: System
**メニュー名**: Restore Point
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（事前冪等性チェックのみ。`-Verified` 未渡し）
**サブスクリプト**: なし（ローカルヘルパー `Test-RestoreRegistryValue` / `Get-ShadowStorageInfo`）

## 目的
Windows のシステムの保護（System Restore）まわりを CSV ベースで一括設定するモジュール。
保護有効化、24 時間制限解除、シャドウコピー容量設定、復元ポイント作成の 4 操作を 1 行 1 操作で
列挙する設計で、キッティング開始前のスナップショット取得などに利用する。クライアント OS
（Windows 10 / 11）専用で、サーバー OS では `Enable-ComputerRestore` 等が使えないため動作不可。

## 入力 (CSV)
`restore_point_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `SettingName`: 操作種別 (`enable_protection` / `remove_24h_limit` / `set_storage_size` / `create_restore_point`)
- `Drive`: 対象ドライブ（例 `C:\`、不要操作は空欄）
- `Description`: 説明（表示・復元ポイント名兼用）
- `Value`: 操作別値（容量 % / RestorePointType 等）
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 保護有効化 (C:\) → 24h 制限解除 → シャドウ容量 5% → fabriq キッティング前
ポイント作成 (MODIFY_SETTINGS) の 4 行が Enabled=1。

## 主要ステップ
1. `restore_point_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. `Test-AdminPrivilege` で管理者権限確認
3. ドライラン表示（操作別に冪等性チェック → `[SKIP]` / `[APPLY]` 表示、
   `vssadmin list shadowstorage` で現容量も併記）
4. 実行確認（AutoPilot 自動 Y）
5. 適用ループ (switch で operation 分岐):
   - `enable_protection`: `DisableSR=0` なら Skip、それ以外は `Enable-ComputerRestore -Drive`
   - `remove_24h_limit`: `SystemRestorePointCreationFrequency=0` なら Skip、それ以外は同値書き込み
   - `set_storage_size`: `vssadmin resize shadowstorage /maxsize=N%`（毎回実行、非冪等）
   - `create_restore_point`: `Checkpoint-Computer -Description -RestorePointType`（非冪等）
6. `New-BatchResult` 返却（`-Verified` は渡さない）

## 注意点・運用メモ
- 管理者権限必須、クライアント OS 限定
- CSV は論理依存順（保護有効化 → 制限解除 → 容量設定 → ポイント作成）の並びを推奨
- `set_storage_size` は vssadmin の制約で冪等チェックなし（毎回呼ぶ）
- `create_restore_point` は OS 側の 24 時間制限で実体としては Skip されるケースあり
  （事前に `remove_24h_limit` を実行する設計）
- `Test-RestoreRegistryValue` は `reg_hklm_config` の `Test-RegistryValueMatch` パターンを
  踏襲したローカルヘルパー、`Get-ShadowStorageInfo` は `ssid_config` の netsh パース流儀を踏襲
- 復元ポイントは Description がそのまま表示名になるため CSV の Description 欄は意味のある日本語可

## 検証
operation 別に挙動が異なる:
- `enable_protection` / `remove_24h_limit`: 適用前にレジストリ読み返しで一致なら Skip と判定。
  これが事実上の事前検証だが、Step 5.5 として独立した読み返しは実装していない
- `set_storage_size`: vssadmin 戻り値（`$LASTEXITCODE`）のみで判定。容量設定の最終値読み返し未実装
- `create_restore_point`: `Checkpoint-Computer` 例外なしを成功として扱う

`New-BatchResult` に `-Verified` を渡していないため履歴の Verified 列は空欄。
実効性は手動で `vssadmin list shadowstorage` / `Get-ComputerRestorePoint` 等で確認する運用。


<!-- ============================================================ -->
# === modules/robocopy_config.md ===
<!-- ============================================================ -->

# robocopy_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: Robocopy
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（ExitCode 判定のみ。実ファイル一致検証なし、`-Verified` 未渡し）
**サブスクリプト**: なし

## 目的
CSV (`robocopy_list.csv`) に列挙したジョブを `robocopy.exe` で順次実行するモジュール。
主要オプション (Recursive / Mirror / CopyACL / SkipOlder) を 0/1 フラグ列で表現してタイポ事故を防ぎ、
`/MT` `/Z` `/XD` などの追加引数は CustomOptions 列で自由記述する設計。
UNC 認証（`net use`）を内蔵しており、SMB 共有を含むコピーも 1 ジョブで完結させられる。
ENC: プレフィックスによる暗号化パスワードに対応。

## 入力 (CSV)
`robocopy_list.csv`
- `Enabled`: 1=有効 / 0=無効
- `ID`: ジョブ識別子
- `Source`, `Destination`: コピー元・先パス（ローカル / UNC）
- `Recursive`: 1=`/E`
- `Mirror`: 1=`/MIR`（Recursive と両立時は Mirror 優先）
- `CopyACL`: 1=`/COPYALL /DCOPY:DAT` / 0=`/COPY:DAT /DCOPY:DAT`
- `SkipOlder`: 1=`/XO`
- `FileFilter`: ファイル名 / ワイルドカード（空欄=全ファイル、スペース区切り複数指定可）
- `CustomOptions`: 追加オプション自由記述
- `AuthUser`, `AuthPass`: UNC 認証情報（任意、ENC: 対応）
- `Description`, `Segment`

## 主要ステップ
1. `robocopy_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. 各ジョブのドライラン表示（オプション展開、AuthUser はマスク）
3. 実行確認（AutoPilot 自動 Y）
4. ジョブループ:
   - 4-1. AuthUser/AuthPass 両方ありなら `net use \\Server\Share /user:...` で接続
   - 4-2. `robocopy Source Destination [FileFilter] /R:3 /W:5 /NP <flags>` 実行（baseline オプション強制付与）
   - 4-3. ExitCode 評価 (0=変更なし / 1=コピー成功 / 2-3=余剰検出 / 4-7=ミスマッチ Warning / 8+=Error)
   - 4-4. finally で `net use /delete` 切断
5. 集計して `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（CopyACL=1 で SeSecurityPrivilege/SeBackupPrivilege が必要）
- baseline オプション `/R:3 /W:5 /NP` は CSV から除外不可（ハングアップ防止のセーフガード）
- パスワードはコンソール / トランスクリプトに一切出力しない（ログ漏洩対策）
- Mirror=1 は Source 不在ファイルを Destination から削除するため、誤 Source で意図せぬ削除事故の
  リスクあり。初回テストは Recursive=1 + SkipOlder=1 を推奨
- Mirror=1 は厳密な冪等ではない（Source 変化に追従するため）
- `FileFilter` は Source をディレクトリ指定のままにしてファイル名はこの列で渡す形式
- net use 失敗時はジョブ Skip して次へ進行（ジョブ間隔離）
- robocopy ExitCode 仕様の Warning レンジ (4-7) は Success 扱い

## 検証
robocopy の ExitCode は「ジョブが成功したか」を返すだけで、実ファイル単位の一致は保証しない。
本モジュールは実ファイル比較や ACL 比較の Step 5.5 を持たず、`New-BatchResult` に `-Verified` を
渡していないため履歴の Verified 列は空欄。手動検証としては `robocopy Source Destination /L /E /NP`
（`/L` でドライラン）や `icacls` での ACL 確認が Guide で案内されている。
将来的に DFS-R や Hash 比較を組み込むには copyfile_config 系の検証除外パターン
（`project_verification_exclusions.md` 参照）と整合を取る必要がある。


<!-- ============================================================ -->
# === modules/scheduled_task_config.md ===
<!-- ============================================================ -->

# scheduled_task_config (Standard)

**カテゴリ**: System
**メニュー名**: Scheduled Task Enable / Scheduled Task Disable
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（事前 State 比較のみ。`-Verified` 未渡し）
**サブスクリプト**:
- `scheduled_task_enable_config.ps1` … タスク有効化
- `scheduled_task_disable_config.ps1` … タスク無効化
- 共有 CSV `task_list.csv`

## 目的
Windows タスクスケジューラのタスクを CSV ベースで一括有効化 / 無効化するモジュール。
2 つのスクリプトが同じ `task_list.csv` を参照し、CSV の `Enabled` 列は「行を処理対象にするか」
であって「タスクを有効化したいか」ではない（アクションは実行スクリプトで決まる）共有 CSV 設計。
冪等で、再実行は Skip 扱いになりエラーにならない。

## 入力 (CSV)
`task_list.csv`
- `Enabled`: 1=処理対象 / 0=スキップ
- `TaskPath`: Task Scheduler フォルダパス（先頭・末尾に `\` 必須、ルートは `\` のみ）
- `TaskName`: タスク名（完全一致）
- `Description`: 表示用説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: `Pre-staged app cleanup`（Enabled=1）/ Schedule Scan / Scheduled Start /
ScheduledDefrag / SilentCleanup（いずれも Enabled=0、参考行）

## 主要ステップ
**[Enable]**
1. `task_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. 前提チェック（`Get-ScheduledTask` cmdlet は Windows 標準のため依存なし）
3. 各タスクの現状表示 (`[Disabled]` / `[Ready]` / `[NOT FOUND]`)
4. 実行確認（AutoPilot 自動 Y）
5. 各タスクをループ: 状態が `Disabled` なら `Enable-ScheduledTask`、それ以外は Skip、
   存在しなければ Fail
6. `New-BatchResult` で集計返却

**[Disable]**
同じ流れで `Disable-ScheduledTask`、`Disabled` なら Skip。

## 注意点・運用メモ
- 管理者権限必須
- `TaskPath` の末尾 `\` 漏れは `Get-ScheduledTask` で見つからず Fail になる頻出ミス。
  Guide でも明示的に注意喚起
- 対象タスクが OS にない場合（OS バージョン差や日本語版 / 英語版差）は Warning 扱いで Fail カウント
- `task_list.csv` は Enable / Disable 両スクリプトが共有するため、誤って同じタスクを
  両方の Order に含めると最後に実行されたほうの状態に収束する（運用注意）
- AutoPilot は `Confirm-ModuleExecution` 経由で確認スキップ

## 検証
独立した Step 5.5 は実装されていないが、ステップ 3 の State 表示自体が事前検証として機能する：
2 回目実行時は全タスクが `Already enabled` / `Already disabled` で Skip と表示されるため、
これが事実上の Post-Apply Verification（同じスクリプトを再実行して全件 Skip になることで
意図状態が確定）。`New-BatchResult` に `-Verified` は渡していないため履歴の Verified 列は空欄。
タスクが見つからない (`[NOT FOUND]`) ケースは「対象 OS にそのタスクが存在しない」可能性が
あるため、Fail でも CSV 側の Enabled=0 化で運用回避するのが推奨。


<!-- ============================================================ -->
# === modules/script_looper.md ===
<!-- ============================================================ -->

# script_looper (Extended)

**カテゴリ**: Scripts
**メニュー名**: Script Looper
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（対象スクリプト側が自身の Verified を返すべきという責務分離。ループ側では読み返しなし）
**サブスクリプト**: なし

## 目的
任意の `.ps1` スクリプトを「条件付きでリトライ・ループ実行」する汎用モジュール。
NIC 設定の DHCP 解放待ち、プリンタドライバ DL の一時失敗、Azure AD Join の完了待ち、外部サービスへの接続不安定などを「時間経過＋リトライで解決」させるためのフレームワーク。`azure_ad_join_check` モジュールはこの looper との組み合わせを主用途として設計されている。

## 入力 (CSV)
`looper_list.csv`:
- **Enabled**: 有効フラグ（必須）
- **ScriptPath**: 対象 `.ps1` パス（必須、fabriq ルート相対パス推奨 / 絶対パス可）
- **MaxRetry**: 最大実行回数（必須、初回含む。1=リトライなし、3=初回+2 回リトライ）
- **IntervalSec**: リトライ前の待機秒数（必須、0 以上）
- **Condition**: `OnError`（Error 時のみリトライ） / `Always`（結果問わず MaxRetry まで実行）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. CSV 読み込み
2. 各エントリで ScriptPath 解決（絶対パスならそのまま、相対なら `Get-Location` 起点で `Join-Path`）+ パラメータ検証（MaxRetry≥1, IntervalSec≥0, Condition∈{OnError,Always}）。結果を `_PathValid` / `_ValidParams` プロパティとして付与
3. ドライラン表示（`[READY]`/`[NOT FOUND]`/`[INVALID]` 色分け）
4. `Confirm-ModuleExecution`
5. 各エントリのリトライループ:
   - 対象スクリプトを `& $scriptPath` で実行
   - **Dual ModuleResult detection**: パイプライン出力から `_IsModuleResult=$true` を持つオブジェクト探索 → なければ `$global:_LastModuleResult` を fallback 参照（kernel の `Invoke-KittingScript` と同じパターン）
   - レガシースクリプト（ModuleResult 返却なし）は例外なし=Success / 例外発生=Error として扱う
   - Condition=OnError なら Error 時のみ次試行、Always は常に次試行（最終試行は必ず終端）
   - `IntervalSec > 0` なら `Start-Sleep`、0 なら即時リトライ
6. 各エントリの最終 `lastStatus` で集計（Error なら failCount、それ以外なら successCount）
7. `New-BatchResult` で全体集計

## 注意点・運用メモ
- 無限ループ防止のため MaxRetry は必須。0 以下や非数値は `[INVALID]` でスキップ
- ScriptPath 不在は `[NOT FOUND]` でスキップ（Error にしない）
- 対象スクリプトは `New-ModuleResult` を返す fabriq 準拠形式が推奨。レガシー対応は fallback 機能だが正確なステータス検出はできない
- 集計ロジック: 全成功→Success / 混在→Partial / 全失敗→Error / 全スキップ→Skipped

## 検証
未実装。Guide.txt 内に明示された設計判断として「対象スクリプトの責務として検証を行うべき（各モジュールが自身の Verified を返すべき）」「ループ側では読み返しを行わない」。`-Verified` 未渡しで Verified 列は空欄。


<!-- ============================================================ -->
# === modules/signout_config.md ===
<!-- ============================================================ -->

# signout_config (Standard)

**カテゴリ**: System
**メニュー名**: Sign-Out
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（プロセス自身が消えるため検証不可、`-Verified` 未渡し）
**サブスクリプト**: なし

## 目的
現在のユーザーセッションを `logoff.exe` 経由でサインアウトするモジュール。
Profile 末尾に配置することで、キッティング作業の終了をユーザーセッションごと閉じる
クリーンアップステップとして機能する。Order=100（最終位置）の指定通り、本モジュールが
走ると後続モジュールはまったく実行されない。

## 入力 (CSV)
設定 CSV なし。

## 主要ステップ
1. 警告バナー表示（"fabriq will be TERMINATED after sign-out"）
2. 実行確認（AutoPilot 時は自動 Y）
3. `New-ModuleResult -Status Success -Message "Sign-out initiated"` を先に確定
   （`$global:_LastModuleResult` フォールバックにより、フレームワークが履歴を捕捉できるように）
4. `Show-Warning` で再度警告
5. `Invoke-CountdownSignout -Seconds 7` で 7 秒カウントダウン後にサインアウト実行

## 注意点・運用メモ
- AutoPilot 対応（無人実行を前提とした運用で重要）。
  メモ `feedback_autopilot_wording.md` の通り「完全無人」ではなくオペレーター立ち会い前提だが、
  Sign-Out は Wizard 終了時のセッション切替として意図的に AutoPilot 自動 Y を許可
- Profile 上の配置位置は必ず最後（Order=100）。途中に置くと後続モジュールが死ぬ
- `logoff.exe` は対話セッションの終了が主用途で、コンソールセッションでは
  fabriq 自体が即座に終了する
- ローカル状態 JSON や履歴ファイルへの書き込みは Step 3 で完了させる設計（プロセス死後の
  書き込みは保証されないため）
- カウントダウン中の Ctrl+C などのキャンセル余地は `Invoke-CountdownSignout` 側の実装による

## 検証
本モジュールに Post-Apply Verification は実装されていない。`logoff.exe` の実行直後に
fabriq プロセス自身が終了するため、モジュール内で読み返しを行う余地が技術的に存在しない。
`New-ModuleResult` 呼び出しに `-Verified` を渡していないため履歴の Verified 列は空欄。
事実上の検証は「サインアウトが起き、再ログオン後に Fabriq.exe が起動していないこと」
で確認する運用。


<!-- ============================================================ -->
# === modules/spi_config.md ===
<!-- ============================================================ -->

# spi_config (Standard)

**カテゴリ**: System
**メニュー名**: SPI Config (SystemParametersInfo)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 意図的に非対応（事前 GET 比較で代替、`project_verification_exclusions` 整合）
**サブスクリプト**:
- `spi_config.ps1`（メイン）
- 配置物: `C:\ProgramData\fabriq\apply_spi.ps1`（Active Setup / Startup Batch 有効時に CSV から動的生成）

## 目的
Win32 `SystemParametersInfo` API 経由で、レジストリ直接書き換えでは安全に制御できない
設定（特に `UserPreferencesMask` のビットフラグ群）を CSV 駆動で変更するモジュール。
視覚効果のオン・オフ、マウス速度、キーリピート速度など SPI でしか正しく扱えない項目を
担当する。Default プロファイルへの反映には Active Setup（HKLM Active Setup 登録 →
新規ユーザー初回ログオン時 cmd 経由で apply_spi.ps1 実行）と Startup Batch（Explorer
起動後トリガー）を併用する belt-and-suspenders 設計。

## 入力 (CSV)
`spi_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `SpiAction`: SET 用 SPI 定数（16 進、例 `0x1043`）
- `ValueMode`: `bool` / `uiParam` / `pvParam`（値の渡し方）
- `Value`: 設定値（bool は 0/1）
- `Description`: 表示用
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 視覚効果 8 項目（アニメーション、影、コンボボックススライド、ヒント、
マウスポインタ影、メニュースライド、フェードアウト、リストボックススクロール）を全て OFF。

## 主要ステップ
1. `spi_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. ドライラン表示（GET action = SET action - 1 で現在値を取得し `[SKIP]` / `[APPLY]` 色分け）
3. 実行確認（AutoPilot 自動 Y）
4. SET ループ:
   - `bool` → `pvParam` に 0/1
   - `uiParam` → `uiParam` に整数
   - `pvParam` → `pvParam` に整数
5. CSV bool 項目から `apply_spi.ps1` を動的生成 → `C:\ProgramData\fabriq\` 配置
6. `Register-FabriqActiveSetup -GUID {fabriq-spi-config}` で HKLM Active Setup 登録
7. `$ENABLE_STARTUP_BATCH=$true` ならランチャー / Startup トリガー配置（`Deploy-FabriqUserSetupLauncher` /
   `Deploy-FabriqStartupTrigger`）→ `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（HKLM Active Setup 書き込み、ProgramData 配下への配置）。
  権限不足時はカレントユーザー適用のみ
- Active Setup 対象は `bool` のみ（`pvParam`/`uiParam` はセッション固有のため除外）。
  Startup Batch は全 ValueMode 対象
- `0x0049` (MinAnimate) は SPI ではなくレジストリ書き込みで制御可能なため `reg_hkcu_config` 側で管理
- `reg_hkcu_config` と Startup Batch / Active Setup の補完機構を共有する設計（同じ
  `apply_*.ps1` ランチャーを `Deploy-FabriqUserSetupLauncher` 経由で配置）
- 初回ログオン時に Explorer が一瞬再起動する副作用あり（タスクバーが瞬間的に消える）
- WPA3SAE のような OS バージョン依存ではないが、SPI 定数は OS バージョンによって挙動差あり

## 検証
意図的に Post-Apply Verification 非対応（`project_verification_exclusions.md` の方針と整合）:
1. SPI は SET 前の GET 比較が事実上の事前検証として機能する
2. SET 直後の GET は API 呼び出し順や Active Setup / Startup Batch の後続適用との競合で
   false PASS / FAIL を返すリスクが高い
3. 新規ユーザーへの反映は Active Setup / Startup Batch に委ねる設計のため、
   カレントセッションだけ読み返してもモジュール全体の効果は保証できない

このため履歴の Verified 列は空欄。SPI GET 比較で不一致が発生した場合は SET 自体を
エラー報告する形で品質を担保する。


<!-- ============================================================ -->
# === modules/ssid_config.md ===
<!-- ============================================================ -->

# ssid_config (Standard)

**カテゴリ**: Network
**メニュー名**: SSID Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（プロファイル一覧再列挙で存在確認）
**サブスクリプト**: なし（XML を実行時に動的生成し %TEMP% 経由で `netsh` に渡す）

## 目的
Wi-Fi (WLAN) プロファイルを `netsh wlan add profile` で登録するモジュール。
CSV に定義した SSID ごとに WLAN プロファイル XML を動的生成し、`%TEMP%` に UUID 付きで
書き出して netsh に取り込ませる。WPA2PSK / WPA3SAE / open の 3 認証方式に対応。
パスフレーズは平文だけでなく `ENC:` プレフィックス付きの暗号化文字列にも対応し、
Fabriq Studio で暗号化したものを実行時に自動復号する。

## 入力 (CSV)
`ssid_list.csv`
- `Enabled`: 1=登録 / 0=スキップ
- `SSID`: ネットワーク名
- `Authentication`: `WPA2PSK` / `WPA3SAE` / `open`
- `Encryption`: `AES` / `none`（open は none）
- `Password`: パスフレーズ（open は空欄、`ENC:` 暗号化対応）
- `AutoConnect`: 1=auto / 0=manual
- `NonBroadcast`: 1=隠れネットワーク / 0=通常
- `Description`, `Segment`

## 主要ステップ
1. `ssid_list.csv` 読み込み（Enabled=1 のみ）
2. 前提チェック: `netsh wlan show profiles` で WLAN サービス応答 + 既存プロファイル名抽出
   （日本語版・英語版両方の出力に対応するパース）
3. ドライラン表示（既存プロファイルと大文字小文字無視で照合 → `[SKIP]` / `[ADD]`）
4. 実行確認（AutoPilot 自動 Y）
5. 適用ループ:
   - 既登録は Skip（冪等性、上書きはしない）
   - Authentication に応じた XML を動的生成（open は sharedKey ブロック無し、
     WPA2PSK/WPA3SAE は passPhrase 平文埋め込み）
   - SSID / Password は `[SecurityElement]::Escape` で XML エスケープ
   - 一時 XML を `%TEMP%\<UUID>.xml` に生成 → `netsh wlan add profile filename=...` 実行
   - finally 句で一時 XML を必ず削除（パスワード残留防止）
6. Step 5.5: Post-Apply Verification（再度 `netsh wlan show profiles` 実行 → 全対象 SSID が
   プロファイル一覧に存在するか検証）→ `New-BatchResult -Verified $verified` 返却

## 注意点・運用メモ
- 管理者権限必須
- WLAN サービス無効・無線アダプタ無しの環境では前提チェックで失敗
- 既存プロファイルは上書きしない設計（パスワード変更等は事前に `netsh wlan delete profile` 必要）
- WPA3SAE は Windows 10 2004 以降 + 対応アダプタのみ機能。非対応環境では登録は成功するが
  接続時にエラーになる可能性
- `ENC:` 復号失敗時は当該行のみエラーで他は継続
- パスワードがコンソール / トランスクリプトに出力されないようマスキング徹底
- 一時 XML は中断時でも finally で削除（パスワード残留リスク最小化）

## 検証
Step 5.5 では `netsh wlan show profiles` を再実行して出力をパースし、
登録対象 SSID が全件プロファイル一覧に含まれているかを大文字小文字無視で照合。
1 件でも欠損があれば `-Verified=$false` で `New-BatchResult` に渡す。
ただし「プロファイルが存在する」ことの検証であり、「実際にその AP に接続できる」ことまでは
保証しない（電波状況や AP 側設定は対象外）。プロファイルの中身（パスワード値等）も読み返し
比較していない（`netsh wlan show profile name=X key=clear` を呼ぶことになりログに平文が
残るため意図的に避けている）。


<!-- ============================================================ -->
# === modules/startlayout_config.md ===
<!-- ============================================================ -->

# startlayout_config (Standard)

**カテゴリ**: Desktop
**メニュー名**: Start Layout Backup / Build / Import / Delete
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（Backup/Delete はファイル・cmdlet レベル検証あり、`-Verified` 未渡し）
**サブスクリプト**:
- `startlayout_backup_config.ps1` … 現スタートレイアウトを `Export-StartLayout` で JSON 出力
- `startlayout_build_config.ps1` … JSON → customizations.xml → .ppkg を `ICD.exe` でビルド
- `startlayout_import_config.ps1` … `Install-ProvisioningPackage` で PPKG 適用
- `startlayout_delete_config.ps1` … 既インストール PPKG をアンインストール + ファイル物理削除

## 目的
Windows 11 のスタートメニュー（ピン留めレイアウト）を「Backup → Build → Import → Delete」
の 4 フェーズで管理するモジュール群。Backup でひな型を抽出し、Build で Provisioning
Package (.ppkg) に変換し、Import で配布、Delete で解除という ppkg ベースの正規ルートを
1 モジュールに収めている。Win11 の `Export-StartLayout` JSON 形式が前提なので
Win10 環境では非対応。

## 入力 (CSV)
`startlayout_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `Id`: 連番識別子
- `FileName`: 拡張子なしの共通ベース名（`startlayout` 指定時に
  `json/startlayout.json` / `xml/startlayout.xml` / `ppkg/startlayout.ppkg` が入出力対象）
- `Segment`: Segment フィルタ（任意）

サブディレクトリ `json/` `xml/` `ppkg/` は実行時自動作成。

## 主要ステップ
**[Backup]** 1. CSV読込 2. `Export-StartLayout` 存在確認 3. ドライラン (`[NEW]`/`[OVERWRITE]`)
4. 確認 5. JSON 出力 6. ファイル存在 + サイズ>0 検証

**[Build]** 1. CSV読込 2. ICD.exe / StoreFile / 入力 JSON 検出 3. ドライラン (`[NEW]`/`[REBUILD]`)
4. 確認 5. JSON 圧縮 + XML エスケープ → `customizations.xml`（UTF-8 BOM）生成 →
`ICD.exe` で .ppkg ビルド

**[Import]** 1. CSV読込 2. `Install-ProvisioningPackage` cmdlet と PPKG ファイル確認
3. ドライラン (`[INSTALL]`/`[REINSTALL]`) 4. 確認 5. `Install-ProvisioningPackage -QuietInstall -ForceInstall`

**[Delete]** 1. CSV読込 2. `Get-/Remove-ProvisioningPackage` 確認 3. ドライラン
(`[INSTALLED]`+PackageId / `[NOT FOUND]`) 4. 確認 5. Phase 1 cmdlet 削除 → Phase 2 .ppkg 物理削除
→ Phase 3 再クエリで完全解除を検証

## 注意点・運用メモ
- Win11 専用（Export-StartLayout JSON 形式が前提）
- Build は Windows ADK の Imaging and Configuration Designer 必須。
  ICD.exe 検索パスは `${ProgramFiles(x86)}\Windows Kits\10\...` 系を 2 候補 + PATH の順に走査
- Import/Delete は管理者権限必須、Backup は不要
- Delete は PPKG 経由で配布した他の設定（ポリシー等）も同時に解除される副作用あり
- Build に渡す JSON は `pinnedList` キー前提（無くても継続するが Warning）
- 典型運用は「Build は手動、Import を Profile に組み込み」（ICD.exe 環境を全クライアントに
  揃えなくて済むように）

## 検証
- **Backup**: 出力後にファイル存在 + サイズ>0 を内部検証
- **Build**: 中間 XML / 最終 PPKG の生成可否は実装で検証するが `-Verified` 未渡し
- **Import**: `Install-ProvisioningPackage` 戻り値のみ。実際にスタートレイアウトに反映されるかは
  再ログオンが必要なケースあり（その場で読み返し検証は不可）
- **Delete**: Phase 3 で `Get-ProvisioningPackage` 再クエリにより完全削除を確認するが
  `-Verified` 引数は未使用

いずれも `New-BatchResult -Verified` 未渡しのため履歴 Verified 列は空欄。
最終的な「ピン留めが意図通り適用された」確認はサインイン後の目視か Get-StartLayout の
出力比較に委ねる設計。


<!-- ============================================================ -->
# === modules/startup_command_config.md ===
<!-- ============================================================ -->

# startup_command_config (Standard)

**カテゴリ**: Scripts
**メニュー名**: Startup Command Deploy
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（将来の新規ユーザー初回ログオンで効果が出るため、`-Verified` 未渡し）
**サブスクリプト**: なし（CSV から `apply_startup_commands.ps1` を動的生成）

## 目的
新規ユーザーの初回サインイン時に一度だけ実行されるコマンド列を、CSV から動的生成して
Default User の Startup フォルダに仕掛けるモジュール。`reg_hkcu_config` / `spi_config` と
共通の Startup ランチャー (`fabriq_user_setup.ps1`) と Startup トリガー
(`FabriqUserSetup.cmd`) を共有し、すべての `apply_*.ps1` が初回ログオン時に同時に走り、
完了後に Explorer 再起動 + Startup トリガー自己削除という belt-and-suspenders 機構の一翼を担う。

## 入力 (CSV)
`startup_command_list.csv`
- `Enabled`: 1=有効 / 0=スキップ
- `Order`: 実行順（昇順数値ソート）
- `Description`: 表示用説明
- `Command`: 実行コマンド文字列（`cmd.exe /c` 経由）
- `Segment`: Segment フィルタ（任意）

例: タイムゾーン設定、フォルダ作成、追加 PowerShell スクリプト起動など。

## 主要ステップ
1. `startup_command_list.csv` を読み込み（Enabled=1）→ Order 昇順ソート
2. ドライラン表示（生成予定ファイル一覧）
3. 実行確認（AutoPilot 自動 Y）
4. CSV から `apply_startup_commands.ps1` を動的生成 → `C:\ProgramData\fabriq\` に配置
5. `Deploy-FabriqUserSetupLauncher` で `fabriq_user_setup.ps1`（共通ランチャー）配置
6. `Deploy-FabriqStartupTrigger` で `FabriqUserSetup.cmd` を Default User Startup に配置
7. `New-BatchResult` 返却（`-Verified` 未渡し）

新規ユーザー初回ログオン時の挙動: Startup → .cmd → フラグチェック（無し）→ PowerShell
起動 → ランチャー → `apply_*.ps1` 群（hkcu / spi / startup_commands）順次実行 →
Explorer 再起動 → `%LOCALAPPDATA%\fabriq\user_setup_done.flag` 作成 → .cmd 自己削除。

## 注意点・運用メモ
- 管理者権限必須（Default User Startup / ProgramData 書き込み）
- コマンドは新規ユーザーコンテキスト（非昇格）で実行されるため、HKLM 書き込みなどは不可
- 各コマンドは独立実行（1 つの失敗が他をブロックしない）
- 実行ログは `%LOCALAPPDATA%\fabriq\startup_commands.log`
- 再実行は冪等（最新 CSV 内容で `apply_*.ps1` を毎回上書き）
- 2 回目以降のログオンは .cmd 削除済 or フラグ検出で即終了するため、軽量
- `reg_hkcu_config` / `spi_config` の Active Setup / Startup Batch と機構を共有しているため、
  他 2 モジュールを使わなくても本モジュール単独で Startup Batch インフラを敷ける

## 検証
本モジュールに Post-Apply Verification は未実装。配備したトリガーが効くのは将来の
新規ユーザー初回ログオン時であり、キッティング時に読み返しても実効性は判定不能。
そのため `New-ModuleResult` / `New-BatchResult` に `-Verified` を渡しておらず、
履歴の Verified 列は空欄。実効性確認は実際に新規ユーザーアカウントを作成して
ログオンし、`%LOCALAPPDATA%\fabriq\startup_commands.log` を確認する手動運用。


<!-- ============================================================ -->
# === modules/storeapp_config.md ===
<!-- ============================================================ -->

# storeapp_config (Standard)

**カテゴリ**: Applications
**メニュー名**: Remove Store Apps
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（AppxPackage / ProvisionedPackage の残存を再チェック）
**サブスクリプト**: なし

## 目的
不要な Microsoft Store / UWP アプリ（Cortana, Bing News, Solitaire, Xbox 系など）を一括削除する
モジュール。「現在のユーザーから削除」(`Remove-AppxPackage`) と「プロビジョニング済み
パッケージ削除」(`Remove-AppxProvisionedPackage`) の両方を行うことで、現ユーザーだけでなく
将来作成される新規ユーザーにも反映されるよう二段階で抹消する。

## 入力 (CSV)
`storeapp_list.csv`
- `No`: 表示順番号
- `AppName`: パッケージ名（例 `Microsoft.BingNews`、`Microsoft.549981C3F5F10`=Cortana 等）
- `Enabled`: 1=削除対象 / 0=残す
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱（既定で削除対象 Enabled=1）: Cortana, Bing News/Weather/Search, Gaming App,
Get Help, Get Started, Office Hub, Solitaire Collection, People, Power Automate Desktop,
Store Purchase, Todos, Alarms, Mail & Calendar, Feedback Hub, Maps, Xbox 系 4 種,
Phone Link, Groove Music, Movies & TV, Outlook for Windows, MSTeams, Quick Assist など多数。

`Get-AppxPackage | Select Name` で実環境のパッケージ名を確認可能。

## 主要ステップ
1. `storeapp_list.csv` 読み込み（Enabled=1 のみ）
2. Appx 関連 cmdlet（`Get-/Remove-AppxPackage` / `Get-/Remove-AppxProvisionedPackage`）の存在確認
3. ドライラン表示（インストール済 / プロビジョニング済の現状を表示）
4. 実行確認（AutoPilot 自動 Y）
5. アプリループ:
   - 5-1. `Get-AppxPackage -Name $AppName -AllUsers` → `Remove-AppxPackage`
   - 5-2. `Get-AppxProvisionedPackage` → `Remove-AppxProvisionedPackage -Online`
6. Step 5.5: 削除後に `Get-AppxPackage` / `Get-AppxProvisionedPackage` で再確認、
   両方残存なし → `[VERIFIED]`、片方でも残れば `[VERIFY FAILED]`
7. `New-BatchResult -Verified $verified` 返却

## 注意点・運用メモ
- 管理者権限必須（ProvisionedPackage 操作のため）
- 一部 OS 内蔵アプリは Remove-AppxPackage を拒否することがあり、その場合は
  `[VERIFY FAILED]` で計上される（OS バージョン依存）
- AppName のワイルドカード非対応（厳密一致）。新製品が出たら CSV 追記が必要
- MSTeams のような新形式パッケージは旧版 (`Microsoft.Teams`) と命名が違うため要注意
- Microsoft.WindowsCalculator のような業務必須アプリは Enabled=0 のままにする運用判断が必要
- 既定の CSV は「業務 PC のキッティング前提でのデフォルト削除セット」として整備されている

## 検証
Step 5.5 でアプリごとに `Get-AppxPackage -Name $AppName -AllUsers` と
`Get-AppxProvisionedPackage -Online` を再実行し、両方の結果が空（残存なし）の場合のみ
`[VERIFIED]` とカウント。1 件でも `[VERIFY FAILED]` があれば `-Verified=$false` で
`New-BatchResult` に渡す。

ただし「現在のユーザーセッション」と「Default プロファイル」の両方の状態は
`-AllUsers` / `-Online` フラグで一括判定しているため、特定ユーザーに残るパッケージは
検出されないケースがある。本モジュールが扱うのはあくまでマシンレベルの
プロビジョニング解除 + 全ユーザー削除であり、既存ユーザープロファイルに残るアプリの
完全クリーンアップは Profile delete 系モジュールに委ねる設計。


<!-- ============================================================ -->
# === modules/sysprep_config.md ===
<!-- ============================================================ -->

# sysprep_config (Standard)

**カテゴリ**: System
**メニュー名**: Sysprep Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 意図的に非対応（`/shutdown` モードで自プロセスが消えるため検証不可）
**サブスクリプト**: なし（CSV から `unattend.xml` / `SetupComplete.cmd` を動的生成）

## 目的
Sysprep の応答ファイル `unattend.xml` と初回起動スクリプト `SetupComplete.cmd` を 3 つの CSV
（`sysprep_list.csv` / `unattend_list.csv` / `setupcomplete_list.csv`）から動的生成・配置し、
最後に `sysprep.exe /generalize` を実行するモジュール。キッティングの「最終一段前」位置に配置し、
これ以降のイメージ封入工程の入口として機能する。第 1 確認（ファイル生成）と
第 2 確認（Sysprep 実行）の 2 段階確認を持ち、第 2 確認 N 選択時はファイル配置済み・未実行で
Status=Success を返して終了する。

## 入力 (CSV)
- `sysprep_list.csv`: `Enabled` / `SysprepExe`（フルパス）/ `Mode` (`oobe`/`audit`) /
  `Shutdown` (`shutdown`/`reboot`/`quit`) / `Description` / `Segment`
- `unattend_list.csv`: `Enabled` / `SettingName` (`ComputerName` / `CopyProfile` / `TestUserName` /
  `EnableAdministrator` / `AdminPassword` / `HideEULAPage` / `ProtectYourPC` /
  `HideWirelessSetupInOOBE` / `HideOnlineAccountScreens` / `HideOEMRegistrationScreen`) /
  `Value` / `Description` / `Segment`
- `setupcomplete_list.csv`: `Enabled` / `Order` / `ActionType` (`DeleteUser` / `CopyFile` /
  `Command`) / `Target` / `Destination` / `Description` / `Segment`
- `source/`: `CopyFile` アクションで使うファイル群を置くディレクトリ

すべて UTF-8 BOM 必須（`feedback_ps1_utf8_bom.md` の方針と整合）。

## 主要ステップ
1. 3 つの CSV を読み込み（sysprep は有効 1 行のみ、unattend はキー単位適用、setupcomplete は
   1 件以上）
2. 前提チェック: `sysprep.exe` 存在 / `source/` 存在 / `C:\Windows\Setup\Scripts\` 自動作成
3. 一覧表示（Sysprep 設定 / Unattend キー / SetupComplete アクション、`[NEW]`/`[OVERWRITE]`）
4. 第 1 確認: ファイル生成・配置（AutoPilot 自動 Y）
5. 配置:
   - 5-1. `source/` → `C:\Windows\Setup\Scripts\source\` ステージング (xcopy)
   - 5-2. `unattend.xml` 動的生成 → `C:\Windows\System32\Sysprep\unattend.xml`
   - 5-3. `SetupComplete.cmd` 動的生成 → `C:\Windows\Setup\Scripts\SetupComplete.cmd`
     （先頭でログを `SetupComplete.log` に追記）
6. 第 2 確認: Sysprep 実行（N 選択時はファイル配置済 + Status=Success で return）
7. `sysprep.exe /generalize /<mode> /<shutdown>` 実行
   （`shutdown`/`reboot` は PC 停止のため Step 7 後は到達せず、`quit` のみ ExitCode で判定）

## 注意点・運用メモ
- 管理者権限必須
- `<LocalAccount>` ブロック（TestUserName 指定時に生成）が unattend にあると Windows 仕様で
  アカウント関連 OOBE 画面が個別設定に関わらず自動スキップされる。OOBE を表示させたい場合は
  TestUserName 行を Enabled=0 にする
- `<AdministratorPassword>` はパスワード設定のみ。Administrator アカウント有効化は
  `setupcomplete_list.csv` の `net user Administrator /active:yes` を Enabled=1 にする
- 冪等性なし（実行のたびに上書き）。Sysprep 自体に同一イメージへの実行回数制限あり（既定 3 回）
- Default プロファイルを露出するための ATTRIB -H / 内部キャッシュ削除 / ATTRIB +H 一連の
  アクションが setupcomplete に標準同梱されている（INetCache / WebCache / LocalLow 削除）
- `taskbar_config` から自動コピーされる `LayoutModification.xml` が source/ 経由で
  Default User の Shell 配下へ展開されるという連携あり

## 検証
意図的に Post-Apply Verification 非対応（`project_verification_exclusions.md` 整合）:
`/shutdown` `/reboot` モードでは sysprep.exe 完了 = PC 停止のためスクリプト側に検証ステップを
走らせる余地がない。`/quit` モードでは検証可能だが通常運用では使わないため統一的に未実装。
`-Verified` を渡していないため履歴 Verified 列は常に空欄。

完了確認は手動で:
- `C:\Windows\System32\Sysprep\Panther\setuperr.log`
- `C:\Windows\Setup\Scripts\SetupComplete.log`
- 配置された `unattend.xml` / `SetupComplete.cmd`


<!-- ============================================================ -->
# === modules/system_finalize.md ===
<!-- ============================================================ -->

# system_finalize (Standard)

**カテゴリ**: Maintenance
**メニュー名**: System Finalize
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（Explorer 再起動 + キャッシュ再生成で false FAIL を招くため意図的に未実装）
**サブスクリプト**: なし

## 目的
キッティング作業の総仕上げとして、シェル環境のリフレッシュを行うモジュール。
shell32.dll の再登録（File Type Association などのデフォルトを再構築）、Explorer 停止、
アイコン / サムネイル / レガシーアイコンの各キャッシュファイル削除、Explorer 再起動を
1 連で実行する。多数の個別設定モジュール適用後に「OS 側のシェルキャッシュが古い」状態を
クリーンアップして反映を確実にする位置づけ。

## 入力 (CSV)
設定 CSV なし（操作内容が完全固定のため）。

## 主要ステップ
1. 対象キャッシュファイルの存在状況を表示（`%LOCALAPPDATA%\Microsoft\Windows\Explorer\`
   配下の `iconcache*.db` / `IconCache*.db` / `thumbcache*.db` / レガシー `IconCache.db` 等）
2. 実行確認（AutoPilot 自動 Y）
3. 順次実行:
   - 3-1. `regsvr32 /s /i:U shell32.dll` で shell32.dll 再登録
   - 3-2. Explorer 停止 (`Stop-Process` / taskkill)
   - 3-3. iconcache*.db / IconCache*.db 削除
   - 3-4. thumbcache*.db 削除
   - 3-5. レガシー IconCache.db 削除
   - 3-6. Explorer 再起動
4. 結果集計表示

## 注意点・運用メモ
- 管理者権限必須（shell32.dll 再登録、システム Explorer の停止 / 再起動に必要）
- 対象キャッシュは現在のユーザー (`%LOCALAPPDATA%`) 配下のみ。他ユーザーは対象外
- 実行中はタスクバーとデスクトップが一時的に消える（Explorer 停止のため）副作用あり
- ロックされたファイルは削除できないことがあり、その場合は Warning 表示で他は継続
- Order=99 は `restart_config` (99) や `signout_config` (100) と並ぶ「最終位置」帯。
  Profile では Sysprep の前 / リスタート前に置いて使うのが想定運用
- ICONCACHE 系の不可視属性 (Hidden) ファイルは Force 指定で削除

## 検証
本モジュールに Post-Apply Verification は実装されていない。
主処理が「Explorer 再起動」であり、Explorer は再起動直後に OS 側でアイコンキャッシュ等を
即座に再生成し始めるため、削除直後にファイル存在チェックを行うと「再生成された新ファイル」
を検出して false FAIL 扱いになるリスクがある。`-Verified` を `New-BatchResult` に
渡していないため履歴の Verified 列は空欄。

実効性確認はオペレーターの目視（タスクバーアイコンが正しく更新されたか、
ピン留めアプリが意図通り表示されるか）に委ねる設計。


<!-- ============================================================ -->
# === modules/taskbar_config.md ===
<!-- ============================================================ -->

# taskbar_config (Standard)

**カテゴリ**: Desktop
**メニュー名**: Taskbar Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（適用先が Default プロファイル＝将来の新規ユーザーのため `-Verified` 未渡し）
**サブスクリプト**: なし（CSV から `LayoutModification.xml` を動的生成）

## 目的
タスクバーにピン留めするアプリを CSV ベースで定義し、`LayoutModification.xml` を
動的生成して Default User プロファイルへ配置するモジュール。配置後に作成される新規
ユーザーがこの XML を継承する形で、最初からタスクバーに業務用アプリのみが並んだ状態で
セッションを開始できる。生成 XML は `sysprep_config/source/` にも自動コピーされ、
Sysprep ワークフローの素材として 1 度の実行で 2 箇所同時に最新化される設計。

## 入力 (CSV)
`taskbar_list.csv`
- `Enabled`: 1=ピン留め対象 / 0=スキップ
- `Order`: タスクバー上の表示順（昇順ソート）
- `LinkPath`: アプリパス（`.lnk` または `.exe`、`%APPDATA%` 等の環境変数使用可）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: File Explorer (Order=10) / Google Chrome (Order=20)。

## 主要ステップ
1. `taskbar_list.csv` 読み込み → Order 昇順ソート
2. 各 LinkPath の存在確認（実行時点で不在でも Warning のみで処理継続）
3. ドライラン表示
4. 実行確認（AutoPilot 自動 Y）
5. `LayoutModification.xml`（CustomTaskbarLayoutCollection / TaskbarPinList）を動的生成 →
   `C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml` 配置
6. 同 XML を `modules\standard\sysprep_config\source\LayoutModification.xml` に自動コピー
   （無条件・毎回。コピー失敗は警告のみで Success 維持）
7. ファイル存在 + サイズ>0 チェック後 `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（Default User プロファイルへの書き込み）
- 既存 `LayoutModification.xml` は上書き
- LinkPath がキッティング時点で不在でも XML には記載される（後続モジュールでアプリを
  入れる Profile でも問題なし）
- 全エントリ Enabled=0 の場合、空の TaskbarPinList を持つ XML が生成され、
  デフォルトのピン留め（Edge 等）も除去される（確信犯的な無効化用途あり）
- `sysprep_config` の `setupcomplete_list.csv` には対応する `CopyFile` 行が同梱済みで、
  Sysprep 実行時に source/ → Setup\Scripts\source\ → Default User Shell\ への二段ホップで
  最終配置される
- ショートカット格納場所:
  - 全ユーザー共通: `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\`
  - 現在ユーザー: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\`

## 検証
Post-Apply Verification は実装されていない。XML 書き出し直後にファイル存在 + サイズ>0 の
最低限チェックは行うが、XML 内容のパース検証ステップはなし。
そもそも実際にタスクバーに反映されるのは「配置後に新規作成される Default User プロファイル」
であり、現セッションでの読み返し検証は仕組み上困難（現ユーザーには既存タスクバーが
存在するため、本モジュールの効果は次のユーザーアカウント作成時まで観測できない）。
`-Verified` を `New-ModuleResult` に渡していないため履歴の Verified 列は空欄。

実効性確認は新規ユーザー作成 → ログオンしてタスクバーを目視、または Sysprep 後のイメージ
展開先での OOBE 後タスクバー確認。


<!-- ============================================================ -->
# === modules/temp_ipaddress_config.md ===
<!-- ============================================================ -->

# temp_ipaddress_config (Standard)

**カテゴリ**: Network
**メニュー名**: Temp IP Address
**VERSION**: 0.1.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（IP/Prefix/Gateway/DNS 読み戻し + Gateway ping）
**サブスクリプト**: なし（GUI 選択ダイアログを WinForms で同居）

## 目的
顧客要件で「本番 IP がまだ旧 PC で使用中」の状況下、新 PC を一時的な作業 IP で立ち上げる
ためのモジュール。CSV に列挙された候補 IP プールを GUI で作業者に提示し、選択された IP を
NIC に付与する。Windows DAD（Duplicate Address Detection）が assignment 時に同 LAN 上の
衝突を検出し、衝突した行は `[DUPLICATE]` でグレーアウト表示して再選択を促す。本番 IP への
切替は別モジュール `ipaddress_config` で行う排他関係。

## 入力 (CSV)
`temp_ipaddress_list.csv`
- `Enabled`: 0=スキップ / 1=プールに含める
- `IPAddress`: 候補 IP（IPv4、strict 4-octet）
- `SubnetPrefix`: CIDR prefix (1-32)
- `Gateway`: デフォルト GW（空欄ならスキップ）
- `DNS1`, `DNS2`, `DNS3`: DNS（空欄なら DNS 設定変更しない）
- `AdapterPattern`: NIC 名ワイルドカード（例 `Ethernet*`）— 全行同一推奨（v1 では混在不許可）
- `Description`: GUI 表示用
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 192.168.100.201〜210（10 IP）すべて Enabled=0 のテンプレート。

## 主要ステップ
1. `Test-AdminPrivilege` チェック
2. CSV 読込（Enabled=1）
3. 各行 validate（IP/Prefix/Gateway/DNS 形式チェック）
4. 全行の AdapterPattern 一致確認 → NIC 解決
5. NIC subnet 整合性チェック（情報表示のみ、機能は継続）
6. GUI ダイアログ表示（AutoPilot でも必ず modal、kitting 中 1 回限りの戦略的判断のため）
7. 選択 IP が現 IP と一致なら sticky SKIP（reassignment なし）
8. assignment（既存 IP/Route 削除 → `New-NetIPAddress` + DNS 設定）
9. DAD 検証（500ms 待機 → AddressState 確認、Duplicate なら roll back + GUI 再表示）
10. Step 5.5: Post-Apply Verification（IP / Prefix / Gateway / DNS 読み戻し + Gateway へ
    `Test-Connection` 1 発）
11. `New-ModuleResult -Verified $verified` 返却

## 注意点・運用メモ
- 管理者権限必須
- AutoPilot 実行中でも GUI ダイアログは必ず表示（自動選択しない）= 設計判断
- 事前 probe（ICMP/ARP）は honest design として行わない:
  「真新しい kitting PC は IP 未取得 → ICMP/ARP 送れない」「切断中 PC は応答せず FREE 誤判定」
  といった限界があるため、信頼できない probe より人間の協調 + DAD + 出荷前検査の三段構えで
  運用カバーする方針
- 既存 IP/Route は assignment 時に削除（DHCP 解除含む）
- 切断中 PC が後で再接続したケースの衝突は DAD 検出不能 = 根本制約
- VERSION は 0.1.0（dev/template ベースで未リリース扱い）

## 検証
Step 5.5 で `Get-NetIPAddress` / `Get-NetRoute` / `Get-DnsClientServerAddress` で読み戻して
CSV 値と比較。AddressState=Preferred を確認。Gateway が指定されていれば `Test-Connection -Count 1`
で疎通を 1 発確認し、reachable なら `-Verified $true` を `New-ModuleResult` に渡す。
ICMP がブロックされている環境では Verified=false になりうるが、IP 設定自体は適用済み。

GUI ダイアログには Status 列に `[CURRENT]`（現 IP と一致）/ `[DUPLICATE]`（DAD 検出済）/
空欄 を表示し、`[CURRENT]` が pool 内にあれば sticky pre-select する UX。


<!-- ============================================================ -->
# === modules/test_error_module.md ===
<!-- ============================================================ -->

# test_error_module (Standard)

**カテゴリ**: Test
**メニュー名**: Test Error Module
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 該当なし（常に Error を返す性質）
**サブスクリプト**: `test_error.ps1`

## 目的
意図的に必ず `Status=Error` で終了するテスト用モジュール。AutoPilot の ErrorMode（`skip` /
`retry` / 未指定時の Retry/Skip ダイアログ）の動作検証に使う。副作用（ファイル / レジストリ /
サービスの変更）は一切なく、どの環境でも安全に走らせられる。本番 Profile に組み込む用途は
想定せず、開発・検証専用の足場部品。

## 入力 (CSV)
設定 CSV なし（module.csv のみで動作）。

## 主要ステップ
1. モジュール表示（"Test Error Module"）
2. 実行確認（Y/N、AutoPilot 自動 Y）
3. `Show-Error` で模擬エラー表示
4. `New-ModuleResult -Status "Error"` を返却

## 注意点・運用メモ
- 管理者権限不要、副作用なし
- 本番 Profile（Master_Config*.csv 等）には絶対組み込まない
- 検証パターン例:
  - **ErrorMode=skip**: モジュール 20 で発生したエラーが自動 skip され 30 が実行される
  - **ErrorMode=retry**: 最大 5 回リトライされ、毎回失敗で最終的に Error として記録
    （AutoPilotMaxRetry=5 が main.ps1 側の上限）
  - **ErrorMode 空欄**: AutoPilot 中に Retry/Skip ダイアログがポップアップ
- メモ `project_autopilot_skip_rejected.md` の通り skip/timeout 機能自体は否決済み
  だが、Profile 列の ErrorMode は別概念で本モジュールはその検証に使う

## 検証
Post-Apply Verification は概念的に該当しない（常に Error を返すモジュールなので
「適用された設定の読み返し」というステップが存在しない）。`-Verified` は渡さず
履歴 Verified 列は空欄。本モジュール自体が他モジュールの Verified 機構の検証ハーネス。


<!-- ============================================================ -->
# === modules/test_harness_config.md ===
<!-- ============================================================ -->

# test_harness_config (Standard)

**カテゴリ**: Test
**メニュー名**: Test Harness
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（CSV から `Verified` 列の値を集計し true/false/null を再現）
**サブスクリプト**: なし

## 目的
fabriq の動作検証用シミュレーションモジュール。ファイルやレジストリに副作用を出さずに、
Status / Verified / AutoPilot ErrorMode / Resume といった全機能パターンを再現できる
動作シミュレーター。Profile 設計時のリハーサルや、main.ps1 / common.ps1 の振る舞い変更時の
回帰テストに使う。Verified の集計、FailFirstN によるリトライ成功シナリオ、cancel Behavior による
Confirm-bypass の Cancelled 再現など、test_error_module よりも表現力が高い。

## 入力 (CSV)
`test_harness_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `Segment`: Profile 側 Segment 値で絞り込むキー
- `TestName`: 表示名 + 状態保存キー
- `Behavior`: `success` / `fail` / `skip` / `cancel`
- `Verified`: `true` / `false` / 空欄（item ごとの Verified 結果）
- `FailFirstN`: 整数。同 (Segment+TestName) の N 回目までは強制 fail
- `DelaySec`: 各 item 処理前のスリープ秒（進捗 UI 確認用）
- `Description`: 説明

デフォルト同梱: success_verified / success_verifail / success_no_verify / skipped / partial
(2 行) / cancelled / error_basic / retry_success (FailFirstN=2) / retry_exhaust /
delay_demo (3s) / hang_sim (120s, `__ASYNC__` + Skip ボタン検証用) / async_ok など。

## 主要ステップ
1. `test_harness_list.csv` 読み込み（Segment フィルタは `Import-ModuleCsv` 共通機構）
2. ドライラン表示（各 item の Behavior / Verified / FailFirstN / DelaySec）
3. 実行確認（AutoPilot 自動 Y）
4. item ループ:
   - DelaySec 指定時はスリープ
   - cancel Behavior は即 Cancelled で全体抜ける（Confirm を介さず確実に Cancelled 再現）
   - FailFirstN 指定 + `$global:FabriqTestHarnessState["${Segment}::${TestName}"]` 累積 ≤ FailFirstN
     なら強制 fail
   - Behavior に応じて Show-Success / Show-Error / Show-Skip と count 更新
5. Verified 集計（全 true→PASS / 1 つでも false→FAIL / 空欄混在→null）
6. `New-BatchResult -Verified $verified` 返却

## 注意点・運用メモ
- 副作用ゼロ（管理者権限不要）。FailFirstN カウンタは `$global:FabriqTestHarnessState`
  ハッシュテーブル上のメモリのみ。`__RESTART__` で PowerShell プロセスが終わると自動リセット
  → ファイル / レジストリ汚染なし
- main.ps1 の `AutoPilotMaxRetry=5` と組み合わせて「N 回失敗 → N+1 で成功」シナリオ再現可能
- cancel Behavior は Confirm を介さないため Stop on Error モード時に Profile 全体停止を
  起こしうる。検証用途以外で使用しない
- 完全なサンプル Profile は `profiles/_test_harness.csv`
- 本番 Profile 厳禁

## 検証
Verified 集計が本モジュールの検証ロジックの中核。CSV の `Verified` 列値を全 item で集約し:
- 1 つでも false → モジュール全体 Verified=FAIL
- すべて true のみ → Verified=PASS
- 空欄混在（false なし） → Verified=null（検証なし扱い）

これは他モジュールの Verified 集計挙動を test_harness 側で意図的に再現できる仕組みで、
common.ps1 の `New-BatchResult -Verified` ロジックをそのまま使う。fabriq 全体の Verified 機構の
動作確認に使う「自作ルーラー」のような立ち位置。


<!-- ============================================================ -->
# === modules/time_sync_config.md ===
<!-- ============================================================ -->

# time_sync_config (Standard)

**カテゴリ**: System
**メニュー名**: Time Sync
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（同期ソース確認、最大 5 回リトライ）
**サブスクリプト**: なし

## 目的
Windows Time サービス (W32Time) に NTP サーバを設定し、時刻同期をトリガーするモジュール。
`w32tm /config /manualpeerlist` で複数サーバを構成し、`/resync /force` で同期を実行した上で、
同期ソースが「Local CMOS Clock」等のローカル系から外れたかを `w32tm /query /status` で
確認する。CSV の Enabled=1 が 0 件のときは「同期のみモード」として、既存の NTP 設定を
保ったまま `/resync /force` のみを実行する 2 モード設計。

## 入力 (CSV)
`time_sync_list.csv`
- `Enabled`: 1=NTP サーバとして使用 / 0=リスト表示のみ（同期のみモードの判定にも使用）
- `NtpServer`: NTP サーバアドレス（例 `time.windows.com`, `ntp.nict.jp`）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: time.windows.com (Enabled=1) / time.google.com (Enabled=0) /
ntp.nict.jp (Enabled=0)。

## 主要ステップ
1. `time_sync_list.csv` 全行読み込み（Enabled フィルタなし、表示と判定の両方に使用）
2. W32Time サービス存在確認
3. ドライラン表示:
   - W32Time 現状 (Status / StartType)
   - 各 NTP サーバへの ICMP 疎通 (`[REACHABLE]` / `[UNREACHABLE]`)
4. 実行確認（AutoPilot 自動 Y）
5. 適用:
   - 5-1. W32Time を Start + StartupType=Automatic
   - 5-2. ICMP 確認（表示のみ、ブロック環境でも継続）
   - 5-3. Enabled=1 の行があれば `w32tm /config /manualpeerlist:"server,0x9 ..." /syncfromflags:manual /reliable:yes /update`
     （`,0x9` = SpecialPollInterval + Client）→ 3 秒待機
   - 5-4. 同期実行（最大 5 回リトライ、3 秒間隔）: `w32tm /resync /force` →
     `w32tm /query /status` でソース確認 → ローカル系以外なら成功
   - 5-5. 最終同期状態を表示
6. Status 判定: ソース切替成功なら Success、5 回失敗なら Partial

## 注意点・運用メモ
- 管理者権限必須
- 標準モード実行時は既存 NTP 設定を上書き
- Hyper-V VM では Time Synchronization Integration Service（ホスト同期）が優先されることがあり、
  w32tm 設定が無視される
- ドメイン参加 PC は通常 DC が NTP ソースになる。意図的に外部 NTP を使う場合のみ本モジュール利用
- ICMP ブロック環境ではドライラン疎通表示が `[UNREACHABLE]` になるが、UDP/123 が通っていれば
  実際の同期は成功する（ICMP は情報表示のみで処理継続）
- 起動直後の同期はサービス初期化遅延で 1-2 回リトライしてから成功するパターンが多い

## 検証
Step 5.4 が事実上の Post-Apply Verification:
- `w32tm /query /status` を再実行して `Source` フィールドを取得
- ソースが「Local CMOS Clock」「LOCL」「Free-running System Clock」のいずれでもなければ
  外部 NTP に切り替わったと判定して成功
- 最大 5 回 (3 秒間隔) リトライして切替えなければ Partial
  （NTP 経路未疎通や起動直後の一時的遅延で発生しうるが、設定自体は正しく適用されているため
  致命的失敗扱いにはしない設計判断）

`-Verified` 引数の渡し方は実装上 Status と連動するため、Success の場合に Verified=true 相当の
履歴が残る。Partial は途中までの同期試行を記録としては残す。


<!-- ============================================================ -->
# === modules/volume_config.md ===
<!-- ============================================================ -->

# volume_config (Standard)

**カテゴリ**: System
**メニュー名**: Volume Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Core Audio API で読み返し）
**サブスクリプト**: なし（C# / COM 経由で `IAudioEndpointVolume` を `Add-Type` で同居）

## 目的
PC 本体のマスター音量とミュート状態を Windows Core Audio API（MMDeviceAPI /
IAudioEndpointVolume）で設定するモジュール。タスクバーの音量スライダーに即時反映され、
再起動不要。受け取り検査時の「全 PC が同じデフォルト音量で出荷される」ことや、
キッティング作業中の意図せぬ大音量再生を抑止する用途で使う。

## 入力 (CSV)
`volume_list.csv`
- `Enabled`: 1=適用 / 0=スキップ
- `Volume`: 0〜100 の整数（%）
- `Mute`: `on` / `off` / 空欄（変更しない）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 1 行 `Volume=50, Mute=off`（マスター音量 50%）。
有効行は最初の 1 行のみ採用される設計（複数行を意味的に重ねない）。

## 主要ステップ
1. `volume_list.csv` 読み込み（最初の Enabled=1 行のみ）
2. オーディオデバイス存在確認（VM でオーディオ無効なら安全終了）
3. `IAudioEndpointVolume::GetMasterVolumeLevelScalar` / `GetMute` で現在値取得・表示
4. 冪等性チェック（既に目標値一致なら Skip）
5. 実行確認（AutoPilot 自動 Y）
6. `SetMasterVolumeLevelScalar` / `SetMute` で即時設定
7. Step 5.5: `GetMasterVolumeLevelScalar` / `GetMute` で読み返し検証 → CSV 期待値と一致確認 →
   `New-ModuleResult -Verified $verified` 返却

## 注意点・運用メモ
- オーディオデバイスが存在しない VM 環境では COM 取得時にエラーになるが安全終了する設計
- デフォルトオーディオデバイスのみ対象（複数デバイスが刺さっている場合は OS 既定デバイス）
- 音量値は scalar (0.0-1.0)、CSV の % 表記は内部で /100 して API に渡す
- Mute=空欄 のときはミュート状態は変更しない（音量のみ変える運用）
- 即時反映なので、キッティング BGM 等を流しているとその場で音量が変わる副作用あり

## 検証
Step 5.5 で `IAudioEndpointVolume::GetMasterVolumeLevelScalar` を再呼び出しして
小数誤差を吸収しつつ ±1% 以内で一致を判定（scalar↔% の往復で丸め誤差が出るため）。
`GetMute` で BOOL を再取得して CSV 値と比較。両方一致で `-Verified=$true`。
即時反映 API なので Verification がそのまま適用結果を反映するパターン。


<!-- ============================================================ -->
# === modules/wallpaper_config.md ===
<!-- ============================================================ -->

# wallpaper_config (Standard)

**カテゴリ**: Desktop
**メニュー名**: Wallpaper Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（API 戻り値判定のみ。レジストリ読み返しなし、`-Verified` 未渡し）
**サブスクリプト**: なし（`SystemParametersInfo` を `Add-Type` で同居）

## 目的
デスクトップ壁紙を画像または単色 (SolidColor) で設定するモジュール。
Windows API (`SystemParametersInfo SPI_SETDESKWALLPAPER`) で即時反映するため再起動不要。
画像壁紙では Style（Fill / Fit / Stretch / Tile / Center / Span）を指定でき、
SolidColor では RGB 値で単色背景を直接適用する。`Resolve-HkcuRoot` 経由で SYSTEM 起動時にも
ログオンユーザーの HKCU ハイブにリダイレクトして書き込む昇格対応。

## 入力 (CSV)
`wallpaper_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `Type`: `Image` / `SolidColor`（省略時 Image）
- `FileName`: 画像ファイル名（`wallpaper/` 配下相対パス）または絶対パス（Type=Image 時）
- `Style`: `Fill` / `Fit` / `Stretch` / `Tile` / `Center` / `Span`（省略時 Fill）
- `Color`: RGB 値スペース区切り（例 `0 99 177`、Type=SolidColor 時）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

サンプル: green.jpg (Image, Fill) / Navy Blue (SolidColor, RGB 0 99 177) いずれも Enabled=0。
画像対応形式: jpg, jpeg, png, bmp, gif, tif, tiff。

## 主要ステップ
1. `wallpaper_list.csv` 読み込み
2. 画像ファイル存在 + 形式確認、または RGB 値妥当性確認 → 一覧表示
3. 実行確認（AutoPilot 自動 Y）
4. Type 別適用:
   - **Image**: `SystemParametersInfo(SPI_SETDESKWALLPAPER, ..., path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)`
     + `Control Panel\Desktop\WallpaperStyle` / `TileWallpaper` レジストリ書き込み
   - **SolidColor**: 壁紙パスをクリア + `Control Panel\Colors\Background` に
     `R G B` 文字列を書き込み + `SystemParametersInfo` トリガー
5. `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（`Resolve-HkcuRoot` で別ユーザーハイブに書き込む可能性のため）
- `wallpaper/` ディレクトリは module ルート配下で、画像をここに置けば相対パス指定可能。
  絶対パスでも OK
- Style 省略時のデフォルトは Fill
- SolidColor 時は FileName / Style は無視
- Sysprep 配布パターンでは Default プロファイルへのコピーは別途必要（本モジュールは
  カレントセッションのみ対象）

## 検証
Post-Apply Verification は実装されていない。
壁紙適用は `SystemParametersInfo` の戻り値（成功 = 非ゼロ）で成否を判定するが、
適用後にレジストリ (`Control Panel\Desktop\WallPaper` / `WallpaperStyle` /
`Colors\Background` 等) を読み返して期待値と一致するかを検証する Step 5.5 はなし。
`-Verified` を `New-BatchResult` に渡していないため履歴 Verified 列は空欄。

理由としては、即時反映 API の戻り値が「OS が壁紙設定を受理した」ことを示すため通常は十分で、
かつレジストリ値（特に Style）の表現が Windows バージョンによって微妙に異なる
（Fill=10/PosV=...）ため一意な比較ロジックを書きにくい、という実装判断。
事実上の確認はオペレータの目視 or `psr.exe` でのスクリーンキャプチャ。


<!-- ============================================================ -->
# === modules/windows_license_config.md ===
<!-- ============================================================ -->

# windows_license_config (Standard)

**カテゴリ**: Security
**メニュー名**: Install Product Key / Activate Windows License
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（Activate 側に再取得確認あり、`-Verified` 引数は未使用）
**サブスクリプト**:
- `windows_license_install.ps1` … プロダクトキー投入
- `windows_license_auth.ps1` … オンライン認証

## 目的
Windows プロダクトキーのインストールとライセンス認証を 2 つのスクリプトで分離して扱うモジュール。
キー投入は `slmgr /ipk` 相当を内部 API で行い、認証は `SoftwareLicensingService.RefreshLicenseStatus`
を起動して結果を `SoftwareLicensingProduct` から再取得することで状態確認する。インターネット未接続
環境では Activate がスキップされる運用設計。

## 入力 (CSV)
`license_key.csv`（Install 側で使用）
- `Enabled`: 1=実行 / 0=スキップ
- `ProductKey`: プロダクトキー (`XXXXX-XXXXX-XXXXX-XXXXX-XXXXX`)
- `Description`: 説明（例: "Windows 11 Pro Volume License"）
- `Segment`: Segment フィルタ（任意）

CSV にキーが無い・無効なら手動入力にフォールバック。Activate 側は CSV を持たない。

## 主要ステップ
**[Install Product Key]**
1. `license_key.csv` からキー読み込み（無ければ手動入力）
2. 現在のライセンス状態を `SoftwareLicensingProduct` から取得・表示
3. 実行確認（AutoPilot 自動 Y）
4. プロダクトキー投入（`SoftwareLicensingService.InstallProductKey` 相当）
5. OS 側のエラーコードで成否判定 → `New-ModuleResult` 返却

**[Activate Windows License]**
1. `SoftwareLicensingProduct` で現状表示
2. 冪等性チェック（`LicenseStatus=1`（Licensed）なら Skip）
3. 実行確認（AutoPilot 自動 Y）
4. `SoftwareLicensingService.RefreshLicenseStatus` 起動
5. 3 秒待機 → `SoftwareLicensingProduct` 再取得 → `LicenseStatus` 確認 →
   1 なら Success、それ以外なら Error

## 注意点・運用メモ
- 管理者権限必須（両スクリプトとも）
- Activate はインターネット接続必須。プロキシ環境では `slmgr.vbs /skms` 等の事前設定が必要
- Volume License (KMS) の場合は KMS サーバ到達性が前提
- License key は CSV 上は平文（Sysprep 用 `unattend.xml` 内の Key も同様）。
  公開リポジトリで誤コミットしないよう運用注意（ENC: 対応は本モジュールでは未実装）
- Order=91/92 で並んでおり、Install → Activate の順で Profile に組み込む想定
- Activate Skip 時のメッセージで「Already Licensed」を明示するため運用ログ上は分かりやすい

## 検証
- **Install Product Key**: Post-Apply Verification は実装していない。プロダクトキーの
  インストール成否は OS 側エラーコードで判定するのみ
- **Activate Windows License**: Step 5 が事実上の Post-Apply Verification。
  `RefreshLicenseStatus` 呼び出し後 3 秒待機 → `SoftwareLicensingProduct` を再取得して
  `LicenseStatus=1` を確認する。ただし `New-ModuleResult` の `-Verified` パラメータは
  現状使用しておらず、履歴の Verified 列は空欄

将来的には `slmgr /xpr` の出力比較 + `LicenseStatus` の数値比較を `-Verified` 統合する
余地あり（`project_crypto_security_review.md` 系の改善対象に含まれる）。


<!-- ============================================================ -->
# === modules/windows_update.md ===
<!-- ============================================================ -->

# windows_update (Standard, Standalone)

**カテゴリ**: なし（standalone module。`module.csv` を持たず、Profile 直接登録不可）
**メニュー名**: メインメニューの `[wu]` ショートカットからのみ起動
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（次回スキャンで 0 件確認が事実上の検証、`-Verified` 未渡し）
**サブスクリプト**:
- `windows_update.ps1` … 1 回分の scan/download/install パス（メイン）
- `wu_launcher.bat` … RunOnce / 再起動後の自動起動用ランチャー
- 状態ファイル: `wu_state.json`（再起動ループ中）/ `wu_completed.json`（完了サマリ、main.ps1 が消費）

## 目的
Windows Update を Microsoft.Update.Session COM API ベースで完全自動化する standalone モジュール。
scan → download → install → reboot → 再 scan のループを自前で持ち、追加更新が無くなるまで
連続適用する。本モジュールは「fabriq の Profile 内モジュール」ではなく独立サイドカー設計で、
ループ制御（RunOnce 登録、状態 JSON ハンドオフ、再起動後再開）は `kernel/main.ps1` の
`Invoke-WindowsUpdateLoop` が担う。fabriq セッション統合のため、完了時には
`wu_completed.json` を書き、次回 main.ps1 起動時に `execution_history.csv` に取り込まれる。

## 入力 (CSV)
`windows_update_list.csv`（キー・バリュー形式）
- `MaxRebootLoops`: 最大再起動回数（既定 5）
- `ScanTimeoutMinutes` / `DownloadTimeoutMinutes` / `InstallTimeoutMinutes`: COM API タイムアウト
- `SuspendBitLocker`: 再起動前 BitLocker 一時停止 (0/1)
- `RebootCountdownSeconds`: 再起動カウントダウン
- `AutoLaunchFabriq`: 完了後 Fabriq.bat 自動起動 (0/1)
- `AutoLogonEnabled`: ワンタイム AutoLogon 設定 (0/1、autologon_list.csv 連動)
- `IncludeOptionalUpdates`: Optional Quality Updates を含める (0/1)
- `IncludeSeekerUpdates`: OptionalInstallation（段階的ロールアウト）を含める (0/1)
- `AutoInstall`: 確認スキップで全更新自動適用 (0/1)

## 主要ステップ
1. ネットワーク疎通 + COM Session 作成
2. `IUpdateSearcher::Search` で更新スキャン（Optional / Seeker フラグに応じて DeploymentAction
   `'Installation' | 'OptionalInstallation'` を加える）
3. 結果一覧表示（Optional は `[Optional]` タグ付き）
4. 確認（`-AutoConfirm` ON / `AutoInstall=1` 時はスキップ）
5. `IUpdateDownloader::Download` でダウンロード
6. `IUpdateInstaller::Install` でインストール
7. 再起動要否判定 → 必要なら `wu_state.json` に保存 + RunOnce(`FabriqWindowsUpdate`) 登録 +
   AutoLogon 設定（CSV ON 時）+ BitLocker 一時停止 → `Invoke-CountdownRestart`
8. 再起動後: `wu_launcher.bat` から本モジュール再起動 → 再 scan → 0 件で完了
9. 完了時: `wu_completed.json` 書き出し → `Invoke-WindowsUpdateLoop` 終了 →
   `AutoLaunchFabriq=1` なら Fabriq.bat 起動

## 注意点・運用メモ
- 管理者権限必須、ネット必須
- `restart_config` と並行使用禁止（両者の RunOnce 登録が競合）
- `MaxRebootLoops`（既定 5）が無限ループ防止セーフティバルブ
- AutoLogon 連動: `windows_update_list.csv` で `AutoLogonEnabled=1` のとき、
  `modules/standard/autologon_config/autologon_list.csv` から `$env:USERNAME` 一致行を引いて
  ワンタイム AutoLogon を設定（不一致なら警告のみで AutoLogon スキップ）
- COM API 操作は同期。ScanTimeout/DownloadTimeout/InstallTimeout が個別保護
- main.ps1 の `Invoke-WindowsUpdateLoop` から呼ぶときに AutoConfirm が立つ。
  AutoInstall=1 を CSV で立てれば main.ps1 を経由しない直接実行でも同じ挙動

## 検証
本モジュールに Post-Apply Verification は実装されていない。
各パスの成否は COM API の install-result HRESULT で判定し、`-Verified` は渡さない。
fabriq の strict な意味での verification は次回 scan で「0 件」が返ることが事実上の検証。
履歴の Verified 列は空欄。

設計判断としては、Windows Update 適用結果の「個々の KB が正しく入ったか」を読み返すのは
WU 自身の差分検出に委ねる方が信頼性が高く（同じ COM API を使うため）、独自 verification を
挟むと重複・矛盾リスクがある、という整理。`wu_completed.json` には
"installed N updates over M reboots" のような要約が入り、`execution_history.csv` 1 行に
畳まれる。


<!-- ============================================================ -->
# === modules/winget_install.md ===
<!-- ============================================================ -->

# winget_install (Standard)

**カテゴリ**: Applications
**メニュー名**: Winget App Installer Update / Winget App Install / Winget App Upgrade
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（ExitCode 判定のみ。`winget list` 再確認なし、`-Verified` 未渡し）
**サブスクリプト**:
- `winget_update.ps1` … winget 自体（Microsoft.AppInstaller）の最新化
- `winget_install.ps1` … `app_list.csv` の未インストールアプリ一括導入
- `winget_upgrade.ps1` … `app_list.csv` の既インストールアプリ一括 upgrade

## 目的
winget (Windows Package Manager) を使ったアプリケーション一括導入・更新モジュール。
3 スクリプトが 1 つの `app_list.csv` を共有し、install と upgrade で「未インストール対象」/
「インストール済対象」を自動振り分けする。winget 自体の更新を別スクリプトに切り出している
のは、古い winget では新しいパッケージリポジトリ形式が読めない問題への対策で、Profile に
組み込むときは 1) winget_update → 2) `__RESTART__` → 3) install/upgrade の順序を推奨。

## 入力 (CSV)
`app_list.csv`（install / upgrade 共有）
- `Enabled`: 1=実行 / 0=スキップ
- `AppID`: winget パッケージ ID（例 `Google.Chrome`、`Adobe.Acrobat.Reader.64-bit`）
- `Options`: 追加オプション（例 `--override "/VERYSILENT"`、空可）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: Google.Chrome / Adobe.Acrobat.Reader.64-bit (Enabled=1) /
Microsoft.VisualStudioCode / Git.Git (Enabled=0)。

## 主要ステップ（共通）
1. `Wait-NetworkReady` でネット接続確認
2. winget (Microsoft.AppInstaller) の利用可否確認
3. `winget source reset --force` で source リセット
4. `app_list.csv` 読み込み → 振り分け（install: 未インストール対象、upgrade: 既インストール対象）
5. ドライラン表示（[INSTALL]/[SKIP インストール済]/[UPGRADE]/[NOT INSTALLED] 色分け）
6. 実行確認（AutoPilot 自動 Y）
7. 順次実行（`winget install/upgrade --id $AppID --exact --silent --accept-source-agreements
   --accept-package-agreements [Options]`）
8. ExitCode 判定 → 結果集計 → `New-BatchResult` 返却

**振り分けマトリクス**:
| 状態 | install | upgrade |
|---|---|---|
| 未インストール | 対象 | スキップ |
| インストール済み | スキップ | 対象 |
| 最新済み | スキップ | Skipped (`-1978335212`) |

## 注意点・運用メモ
- ネット必須、管理者権限必須
- winget の `winget list --id <AppID> --exact` の出力に AppID が現れるかで冪等性判定
- ExitCode 特殊値:
  - `0` 成功 / `3010` 成功（再起動保留）
  - `-1978335212` NO_APPLICATIONS_FOUND → Skipped 扱い（upgrade で「既に最新」）
  - `-1978335189` UPDATE_NOT_APPLICABLE → Skipped 扱い
- `Options` 列でアプリごとのサイレントインストール引数を渡す（一部アプリは winget の
  既定 silent では完全無人にならないため）
- Profile 例: 初回キッティングは update→reboot→install、定期メンテは update→reboot→upgrade、
  フル更新は update→reboot→install→upgrade

## 検証
本モジュールに Post-Apply Verification は実装されていない。
install/upgrade 後の `winget list` 再確認は行わず、ExitCode のみで Success/Failed/Skipped を
決定する。`New-BatchResult` / `New-ModuleResult` に `-Verified` を渡しておらず履歴 Verified 列は
空欄。

理由としては:
1. winget の ExitCode 体系が成功/失敗/「既に最新」を明確に区別する設計のため、
   ExitCode 判定で実用上十分
2. 検証目的の `winget list` 再呼び出しはネットワーク往復を含むため遅く、
   バッチサイズが大きい Profile では実時間に影響する

手動再確認は `winget list --id <AppID>` で可能。

