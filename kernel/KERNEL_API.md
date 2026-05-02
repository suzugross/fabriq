# Fabriq Kernel Public API

**Current Kernel Version**: `3.1.8`（`kernel/KERNEL_VERSION` を真のソースとする）

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
| `$global:AutoConfirmMode` | `bool` | FrexProfile 単発実行中か（`Confirm-Execution` / `Wait-KeyPress` を短絡。AutoPilot のサブセット動作） |
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
| `__ASYNC__` | 以降のモジュールを監視付き Runspace で実行（Skip ボタン / timeout で強制中断可能、since kernel 2.1.0） |
| `__RESTART__` | Windows 再起動 + RunOnce 経由で再開 |
| `__REEXPLORER__` | Explorer 再起動 |
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
- `Register-FabriqRunOnce` / `Invoke-CountdownRestart`
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

---

## 8. API Version History

モジュールの `REQUIRES_KERNEL` 判定（CLAUDE.md ルール I）用の導入バージョン追跡表。各公開 API がどの `KERNEL_VERSION` で利用可能になったかを記録します。

### 2.0.0（baseline）

formal SemVer の出発点。以下すべて利用可能:

- **§1 公開関数**: `Show-Info`, `Show-Success`, `Show-Warning`, `Show-Error`, `Show-Skip`, `Show-Separator`, `Show-CategorySeparator`, `Import-ModuleCsv`, `New-ModuleResult`, `New-BatchResult`, `Confirm-ModuleExecution`, `Confirm-Execution`, `Wait-KeyPress`, `Wait-NetworkReady`, `Test-AdminPrivilege`, `Unprotect-FabriqValue`
- **§2 公開グローバル**: `$global:FabriqMasterPassphrase`, `$global:AutoPilotMode`, `$global:AutoPilotWaitSec`, `$global:FabriqEvidenceBasePath`
- **§3 公開環境変数**: `SELECTED_*` 全般, `SELECTED_PRINTER_<N>_*`, `FABRIQ_SEGMENT`, `FABRIQ_AUTOLOGON_USER`, `FABRIQ_WORKER_NAME`, `FABRIQ_EVIDENCE_BASE`
- **§4 Profile CSV スキーマ**: 列（`Order`, `ScriptPath`, `Enabled`, `Description`, `Segment`, `ErrorMode`）、特殊マーカー（`__AUTOPILOT__`, `__RESTART__`, `__REEXPLORER__`, `__AUTO_to_<User>__`）。
  ※ 2.0.0 baseline には `__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__` も含まれていたが、**3.0.0 で削除**（下記 3.0.0 参照）
- **§5 ModuleResult 契約**: 全フィールド（`Status`, `Message`, `Details`, `Verified`, `Timestamp`, `_IsModuleResult`）

### 2.1.0

- **§4 特殊マーカーに `__ASYNC__` 追加**（プロファイル内で利用するのみ。モジュールスクリプト側から呼び出す API ではないため、`__ASYNC__` を使うプロファイル作者のみこの版を要求する）

### 2.2.0

- **§9「更新・オーバーレイ契約（外部ツール向け公開契約）」新設**
  - `dev/framework_overlay_rules.json`（schemaVersion 1）を単一真実源として明文化
  - bundle 定義（kernel / module）、除外ルール、SemVer 比較セマンティクス、`REQUIRES_KERNEL` 事前チェック、schemaVersion 後方互換を公開契約として固定
  - 本契約は外部更新ツール（fabriq_studio 等）が consume する前提。モジュールスクリプト側の公開 API（§1〜§5）には影響なし。`__ASYNC__` と同様、この版を要求するのは本契約を使う外部ツール側のみ
- **モジュール `VERSION` / `REQUIRES_KERNEL` の baseline 一斉 seed 済み**（本体モジュール 73 件すべてに配備）。ルール H/J が lazy seed から baseline seed 運用へ移行

### 2.2.2

- **§10「Evidence Manifest 契約（外部 evidence consumer 向け公開契約）」新設**
  - `kernel/EVIDENCE_MANIFEST.md`（schemaVersion 1）を単一真実源として明文化
  - `evidence_config` モジュール v1.3.0 以降が `pc_information/<dir>/manifest.json` を出力
  - manifest schema（schemaVersion / sections / status enum / summary 等)、status セマンティクス、前方互換ルールを公開契約として固定
  - 本契約は外部 evidence consumer ツール（fabriq_evidence_manager 等）が consume する前提。モジュールスクリプト側の公開 API（§1〜§5）には影響なし。この版を要求するのは本契約を使うツール側のみ

### 3.0.0

- **§4.2 特殊マーカー 4 種を削除（破壊的変更 / MAJOR）**: `__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__`
  - 削除理由: 実運用での参照ゼロ（`__PAUSE__` / `__STOPLOG__` / `__STARTLOG__`）または唯一の使用箇所も廃止済み（`__SHUTDOWN__`）。fabriq_studio のマーカーパレットでも既に除外されており、UX 上は事実上 deprecated だった
  - 既存プロファイル互換: 削除後のマーカーを含む旧プロファイルは `Resolve-ProfileModules` の `$invalidPaths` 経由で「module not found」warning として降格、kernel はクラッシュせず他モジュールの実行を継続する（graceful degradation）
  - 残存特殊マーカー: `__AUTOPILOT__` / `__ASYNC__` / `__RESTART__` / `__REEXPLORER__` / `__AUTO_to_<User>__` の 5 種
- §6 内部 API 一覧から `Invoke-CountdownShutdown` を削除（`__SHUTDOWN__` 削除に伴うデッドコード除去）

### 3.1.0

- **§2 公開グローバル変数に `$global:AutoConfirmMode` 追加**
  - FrexProfile dashboard の単発実行（`[Run This]`）で `Confirm-Execution` /
    `Wait-KeyPress` を短絡し、Y/N プロンプトと Press-Enter 待機をスキップする
    flag。AutoPilot のサブセット動作で、AutoPilot の inter-module wait /
    ErrorMode 分岐 / `Show-AutoPilotErrorDialog` は発火しない
  - 通常モジュールスクリプトは本グローバルを参照しない（fabriq 本体が
    FrexProfile sub-loop 内でのみ立てる、モジュールから見れば読み取り
    専用 flag）。本値を読むモジュールスクリプトを書く場合のみこの版を要求する
- FrexProfile 機能群の追加: profile CSV を state-aware に部分実行できる
  GUI（`apps/fabriq_operator/lib/frex_dashboard.ps1`）と関連 sub-loop。
  公開 API としてはモジュール側に影響なし — Profile CSV スキーマも特殊
  マーカーも不変。Linear `Execute Profile` 経路は並走運用で従来通り（FrexProfile
  安定後に Linear 撤去予定）
- 内部実装整理:
  - `resume_state.json` に optional `schemaVersion=2` フィールドを追加
    （FrexProfile resume が `SelectedOrders` / `ModuleStates` を保持）。
    Linear が書いた v1 ファイルは引き続き読み書き両方で動作（後方互換）
  - `Complete-ProfileExecution` 関数で post-profile pipeline を集約
    （`Invoke-BatchExecution` 末尾と `[cl]` 再生成の重複を解消）
  - `Invoke-FrexProfileLoop` ヘルパーで FrexProfile sub-loop を一元化
    （main loop "FrexProfile" action と Frex resume bootstrap の単一の
    真実の源）

### 判定ルール

モジュールの `Min Kernel API` は、そのモジュールが使用している公開 API の「導入バージョンの最大値」。

例:
- `Show-Info` + `Import-ModuleCsv` + `New-ModuleResult` のみ使用 → Min Kernel API = **2.0.0**
- 上記に加えて 2.2.0 で追加される新 API を使用 → Min Kernel API = **2.2.0**

プロファイル側（特殊マーカー）の依存はモジュールの `REQUIRES_KERNEL` には含めない（マーカーは kernel が解釈するため、モジュールスクリプト単体の動作には影響しない）。

---

## 9. 更新・オーバーレイ契約（外部ツール向け公開契約）

fabriq 本体の再配布・in-place 更新（site-specific データを保持したままフレームワーク側だけ差し替える）を行う **外部ツール**（代表: fabriq_studio、および `dev/build_framework_patch.ps1`）が依存する契約を本節で明文化します。

### 9.1 真実の源: `dev/framework_overlay_rules.json`

更新ルールは本 JSON に集約されます。schemaVersion 付き。外部ツールは必ずこのファイルを読んでルールを解釈してください。フィールド仕様は JSON 内 `description` および以下。

| キー | 型 | 意味 |
|---|---|---|
| `schemaVersion` | int | マニフェストのスキーマ版。現行 `1`。破壊的変更時に 2 へ |
| `excludeDirsTopLevel` | string[] | robocopy レベルで除外するトップレベルディレクトリ（`.git`, `.claude`, `evidence`, `logs`） |
| `excludeDirsRecursive` | string[] | 再帰的に丸ごと除外するディレクトリ（`profiles` が該当）。**overlay 対象から常時外す** |
| `excludeFilesKernelLevel` | string[] | カーネル配下の個別除外ファイル（`hostlist.csv` / `workers.csv` / `log_destinations.csv` / runtime artifact 全般 / `passphrase_verify.txt` 等） |
| `moduleCsvWhitelist` | string[] | `modules/` 配下の CSV のうち framework 扱いとするホワイトリスト（`module.csv`, `preset.csv` のみ） |
| `bundles.kernel` | object | カーネル bundle の定義（version source + include paths） |
| `bundles.module` | object | モジュール bundle の定義（パターン） |

### 9.2 Bundle 定義

| Bundle | Version ファイル | 対象パス |
|---|---|---|
| **kernel** | `kernel/KERNEL_VERSION` | `kernel/`, `apps/`, `commands/`, `dev/`, `Fabriq.exe`, `Deploy.bat`, `README.md`, `CHANGELOG.md`, `CLAUDE.md`, `LICENSE` |
| **module:\<name\>** | `modules/{std,ext}/<name>/VERSION` | `modules/{std,ext}/<name>/`（ただし `moduleCsvWhitelist` 以外の CSV は除く） |

`apps/` / `commands/` / `dev/` は個別 `VERSION` を持たず、kernel bundle と同期して動きます。

### 9.3 site-specific の絶対保護

以下は **更新時に絶対に上書きしない**：

- `profiles/` 配下全ファイル（`Master_*.csv`, `Custom Plan.csv`, `sysprep.csv`, `_test_harness*.csv`, `easy_template/` 等すべて。プロファイル書式のアップデートが入っても既存を優先）
- `excludeFilesKernelLevel` に列挙された kernel 配下ファイル
- `modules/**/*.csv` のうち `moduleCsvWhitelist` 以外のもの（`_list.csv` ファミリ、`office_key.csv`, `license_key.csv`, `domain.csv` 等）
- ランタイム成果物（`kernel/json/*.json`, `art_pulse.txt`, `skip_request.flag`, `passphrase_verify.txt`, `silence.flag` 等）

### 9.4 SemVer 比較セマンティクス

bundle 単位でバージョンを比較：

| template VERSION | target VERSION | 期待動作 |
|---|---|---|
| `1.2.0` | `1.1.0` | **UPDATE**（overlay 実行） |
| `1.2.0` | `1.2.0` | SKIP（同版） |
| `1.2.0` | `1.3.0` | SKIP（target 側が新しい。ツールによっては警告表示推奨） |
| `1.2.0` | VERSION ファイル欠損 | UPDATE（target を lazy seed） |
| なし | `1.0.0` | SKIP（template 側に VERSION 未打刻） |
| なし | なし | SKIP |

バージョン文字列は `^(\d+)\.(\d+)\.(\d+)$` 形式。pre-release / build metadata は現行不使用。

### 9.5 `REQUIRES_KERNEL` 事前チェック

モジュール bundle を overlay する前に、template 側のそのモジュールの `REQUIRES_KERNEL` と target 側の現行 `kernel/KERNEL_VERSION` を比較：

- `REQUIRES_KERNEL > 現行 kernel` の場合、**先に kernel bundle を overlay する** か、当該モジュール更新を block してユーザに kernel 更新を促す
- この順序を守らないと、新モジュールが古いカーネルの未提供 API を呼んで実行時エラーになる

### 9.6 新モジュール / 欠損モジュールの扱い

- **template にあり target にないモジュール**: overlay で追加（`_list.csv` は copied されない。operator が Fabriq Studio で設定 CSV を新規作成）
- **target にあり template にないモジュール**: site-custom モジュールと見なし **保持**（touched しない）

### 9.7 更新前の安全チェック（推奨）

外部更新ツールが overlay 実行前に必ず確認すべき状態：

| チェック | 目的 |
|---|---|
| `Fabriq.exe` プロセスが実行中でないこと | ファイルロック回避 |
| `kernel/json/resume_state.json` 不在 | キッティング中断中の更新を避ける |
| target の `kernel/KERNEL_VERSION` と全 module の `REQUIRES_KERNEL` の整合 | 更新後のランタイム互換性保証 |
| target フォルダのバックアップ取得 | ロールバック経路確保 |

### 9.8 schemaVersion の後方互換

`dev/framework_overlay_rules.json` の `schemaVersion` フィールドが将来 `2` 等に上がった場合、外部ツールは未対応バージョンを検知したら**処理を拒否して明示エラー**を返す責任があります（黙って部分動作しない）。現行 `1` のスキーマは下位互換を維持する形で進化させる方針。

---

## 10. Evidence Manifest 契約（外部 evidence consumer 向け公開契約）

`evidence_config` モジュールが収集後に出力する `manifest.json` を、外部 evidence consumer ツール（代表: `fabriq_evidence_manager`）が前方互換に消費するための契約を本節で明文化します。

### 10.1 真実の源: `kernel/EVIDENCE_MANIFEST.md`

manifest スキーマ・status セマンティクス・前方互換ルールはすべて `kernel/EVIDENCE_MANIFEST.md` に集約されます。schemaVersion 付き。外部ツールは必ず本ドキュメントを参照して manifest を解釈してください。

### 10.2 配置

```
{evidenceBaseDir}/pc_information/{collectionDir}/manifest.json
```

1 evidence_config 実行 = 1 manifest.json。再実行時は `manifest.json.bak` に rotate（1 世代保持）。

### 10.3 schemaVersion=1 の必須フィールド

| フィールド | 型 | 用途 |
|---|---|---|
| `schemaVersion` | int | 現行 `1`、破壊的変更時に `2` |
| `manifestType` | string | 固定 `"fabriq-evidence-manifest"` |
| `evidenceConfigVersion` | string | manifest を書いた evidence_config モジュールの SemVer |
| `fabriqKernelVersion` | string | manifest 書き込み時点の `kernel/KERNEL_VERSION` |
| `collectedAt` | string (ISO 8601) | 収集開始日時 |
| `computerName` / `hardwareUniqueId` / `selectedNewPcName` | string | PC 識別 3 点組 |
| `sections[]` | array | `{ id, title, files, status, reason, elapsedMs }` |
| `summary` | object | `{ sectionCount, successCount, skippedCount, failedCount, partialCount }` |

### 10.4 status enum

`Success` / `Skipped` / `Failed` / `Partial` の 4 値。詳細は `EVIDENCE_MANIFEST.md` §3.3。

### 10.5 schemaVersion 後方互換

- 未知 major 版を検知した外部ツールは **legacy mode（manifest 無視 + ファイル列挙）にフォールバック** すること。silent な部分動作は禁止
- 未知 section ID は **raw 表示**してクラッシュさせない
- 未知 status 値は **Failed 扱い** で安全側に倒す
- フィールド追加（後方互換）は schemaVersion=1 内で許容、削除・改名・型変更は schemaVersion 昇格を伴う

### 10.6 manifest 不在の旧 evidence

manifest.json 不在の旧形式 evidence（kernel 2.2.1 以前 / evidence_config 1.2.0 以前で収集されたもの）も外部ツールはサポートし続けることが期待されます。manifest が無ければファイル列挙ベースで動作する従来挙動を維持してください。

### 10.7 詳細仕様

完全なスキーマ、サンプル、ディレクトリ表現（`files[]` の `/` 末尾規則）、Partial の使い方等は `kernel/EVIDENCE_MANIFEST.md` を参照。

---

## 11. 変更履歴

公開 API の追加・変更は同一コミット内で本ドキュメントに反映され、`KERNEL_VERSION` の昇格と同期します（CLAUDE.md ルール G）。詳細は `CHANGELOG.md` を参照。
