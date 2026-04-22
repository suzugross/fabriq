# 開発コンテキスト: Windowsキッティングフレームワーク「fabriq」

あなたは現在、Windowsキッティング自動化フレームワーク「fabriq」の開発を行っています。
このフレームワークは、保守性・堅牢性・柔軟性を最優先事項として設計されています。

## 絶対遵守事項（クリティカル・ルール）

これからの作業において、以下のルールを絶対に守ってください。

### 1. 標準テンプレートの厳守 (`dev/template`)

- 新規モジュールやスクリプトを作成する際は、必ず `dev/template` ディレクトリ配下のテンプレートファイルをベースに実装を開始してください。
- テンプレートに含まれる「処理開始ログ」「冪等性チェック」「設定適用」「エラーハンドリング」「終了ログ」の構造フローを勝手に変更しないでください。

### 2. `kernel/common.ps1` の徹底活用

- 独自のログ出力（`Write-Host` 等）や、独自のエラーハンドリング処理を記述することを禁止します。
- 必ず `kernel/common.ps1` を読み込み、そこで定義されている共通関数（`Show-Info`, `Show-Error`, `Show-Success`, `Initialize-Module` 等）を使用してください。
- **関数を新規実装する前に、必ず `kernel/common.ps1` の既存関数一覧を確認してください。** 同等の機能を持つ関数が既に存在する場合、独自関数の作成は禁止です（例: 管理者権限チェックには `Test-AdminPrivilege` を使用）。
- モジュール内にローカルヘルパー関数を定義する場合は、common.ps1 に該当する関数が存在しないことを確認した上で、既存モジュールの類似実装パターン（例: `reg_hklm_config` の `Test-RegistryValueMatch`）を参考にしてください。

### 3. 既存パターンの踏襲（車輪の再発明禁止）

- ロジックの実装で迷った際は、自己判断せず、既存の標準モジュール（`modules/standard/` 配下）の実装パターンを参照・模倣してください。
- 「fabriqらしい書き方」から逸脱しないことが、高機能なコードを書くことよりも重要です。

### 4. 段階的な実装と検証

- いきなりコード全体を書き始めないでください。
- まず「どのテンプレートを使用し、どのようなロジックを組み込むか」を概要レベルで提示し、私の承認を得てからコーディングへ移ってください。

### 5. 検証方法の提示

- 実装コードと共に、その機能が正しく動作しているか、または既存機能を破壊していないかを確認するための具体的な「検証手順（テスト方法）」を必ず出力してください。

### 6. Post-Apply Verification（適用後検証）の推奨

- 設定を適用するモジュールでは、適用後に**システム状態を読み返して期待値と一致するか検証する**ステップ（Step 5.5）の実装を推奨します。
- 検証が実装可能な場合、`New-ModuleResult` / `New-BatchResult` の `-Verified` パラメータを使用して検証結果を返却してください。
- 検証ロジックは、可能な限り既存の冪等性チェック関数（例: `reg_hklm_config` の `Test-RegistryValueMatch`）を再利用してください。
- 再起動後に反映される設定（ホスト名変更等）は、レジストリの保留値など「適用が受理されたこと」を検証してください。
- 検証が技術的に困難なモジュール（例: `sysprep_config`）では `-Verified` を省略（`$null`）して構いません。

## バージョン管理ルール

本プロジェクトはカーネル API とモジュールを独立に SemVer 管理します。
Claude が実装を一手に担う前提で、ランタイムチェックは行わず、**Claude のレビュー手順（実装前宣言 + 実装サマリ）で厳密に制御**します。
[Semantic Versioning 2.0.0](https://semver.org/lang/ja/) 準拠、変更履歴は [`CHANGELOG.md`](CHANGELOG.md)（[Keep a Changelog 1.1.0](https://keepachangelog.com/ja/1.1.0/)）。

### 管理対象

| ファイル | 意味 | 更新タイミング |
|---|---|---|
| `kernel/KERNEL_VERSION` | カーネル API バージョン（真のソース） | カーネル関連変更時。判定基準は `kernel/KERNEL_API.md` の公開 API に影響があるか |
| `kernel/KERNEL_API.md` | 公開 API サーフェスの明文化 | 公開 API の追加・削除・シグネチャ変更時（`KERNEL_VERSION` 昇格と同コミット） |
| `modules/{std,ext}/<name>/VERSION` | モジュール個別バージョン（任意、欠損可） | Claude が初めて touched した時に `1.0.0` を打刻、以降 SemVer で昇格 |
| `dev/template/VERSION` | 新規モジュール用テンプレ（初版 `0.1.0`） | テンプレから作成されるたびに引き継がれる |

全体を表す「ディストリビューション版」ファイルは持ちません。`README.md` L1 / `kernel/common.ps1` L2 / `kernel/main.ps1` L3 の版表記は `KERNEL_VERSION` の `X.Y` に同期します。

### A. コード変更時の義務（毎回）

`kernel/` / `modules/` / `apps/` / `commands/` / `profiles/` / `dev/template/` 配下のコードまたは CSV スキーマを変更した場合、**同じコミット内で必ず** 以下を実施:

1. `CHANGELOG.md` の `[Unreleased]` セクションに該当項目を追記
2. カテゴリ（`Added` / `Changed` / `Deprecated` / `Removed` / `Fixed` / `Security`）を選ぶ
3. 行頭にコンポーネント名（`kernel/common.ps1:`, `modules/standard/<name>:`, `profiles:` 等）のプレフィックス
4. モジュールを更新した場合は該当モジュールの `VERSION` ファイルを作成 or 昇格（欠損時は初版 `1.0.0`、以降 SemVer）
5. 公開 API（`kernel/KERNEL_API.md` の記載範囲）に影響がある場合は、同コミット内で `KERNEL_API.md` を更新

ドキュメント（README, Guide.txt, コメントのみ）の修正は追記不要。

### B. バージョン番号の影響判定

#### カーネル（`kernel/KERNEL_VERSION`）

| 影響 | 昇格 | 例 |
|---|---|---|
| **MAJOR** (X+1.0.0) | 公開 API 破壊的変更 | `KERNEL_API.md` 記載の関数削除・シグネチャ変更 / Profile CSV 必須列削除・改名 / `ModuleResult` フィールド削除・契約変更 / `SELECTED_*` 環境変数改名 |
| **MINOR** (X.Y+1.0) | 公開 API への後方互換な追加 | 公開関数追加 / Profile CSV 任意列追加 / 特殊マーカー追加（例: `__ASYNC__`） / 新環境変数追加 |
| **PATCH** (X.Y.Z+1) | 内部実装のみの変更（公開 API 不変） | `Invoke-SafeCommand` 内部最適化 / `Resolve-ProfileModules` リファクタ / 状態 JSON スキーマ変更 / バグ修正 |

#### モジュール（`modules/{std,ext}/<name>/VERSION`）

| 影響 | 昇格 | 例 |
|---|---|---|
| **MAJOR** (X+1.0.0) | モジュール外部仕様破壊 | `_list.csv` 必須列削除 / モジュールの入出力契約変更 |
| **MINOR** (X.Y+1.0) | 後方互換な機能追加 | 新しい設定項目対応 / 新セグメント追加 / Post-Apply Verification 追加 |
| **PATCH** (X.Y.Z+1) | バグ修正・内部改良 | エッジケース修正 / ログ文言改善 |

判定に迷ったら**大きい側に倒す**。

### C. ルール E: 実装前の事前宣言（必須）

カーネルまたはモジュールを修正する前に、**実装開始前のメッセージで** 以下を宣言する。宣言なしでコード編集に入ることは禁止:

```
【変更スコープ宣言】
- 対象: kernel / module:<name> / profile / doc
- 公開 API サーフェスへの影響: あり / なし
  （あり の場合: どの関数/変数/マーカー/スキーマ が変化するか）
- KERNEL_API.md 参照済み: yes
- 予想バージョン影響:
    kernel  : X.Y.Z → X.Y.Z（MAJOR / MINOR / PATCH / 変更なし）
    modules : <list of touched modules, each with predicted bump>
- 既存モジュールへの波及: ゼロ / <具体リスト>
```

### D. ルール F: 実装サマリでの最終報告（必須）

実装完了報告に以下を含める:

```
【バージョン影響サマリ】
- kernel/KERNEL_VERSION : X.Y.Z → X.Y.Z+N（種別 / 理由）
- KERNEL_API.md の更新 : あり / なし
- touched modules :
    <module_name> : X.Y.Z → X.Y.Z+N（種別 / 理由）
- untouched modules : N/73（一切触っていないモジュール数）
- 配備方針 : kernel/ フォルダ差し替えのみで OK / モジュール X の更新も必要 / 全件再配布必要
```

### E. ルール G: `KERNEL_API.md` の同期保守

公開 API に変更を加える場合、**同じコミット内で** `KERNEL_API.md` を更新する。この更新抜きの `KERNEL_VERSION` MINOR/MAJOR 昇格は禁止。Claude は公開 API を変更する前に必ず `KERNEL_API.md` を読み、現状の宣言内容を把握してから作業に入る。

### F. ルール H: モジュール `VERSION` ファイル運用

- 既存モジュールの `VERSION` ファイルは存在しないことを許容する（欠損＝未追跡扱い）
- Claude が初めて修正を加えた時点で `1.0.0` を打刻する（欠損ファイルの新規作成）
- 以降の修正ごとに SemVer 規則（B 節下段）で昇格
- 新規モジュールは `dev/template/` の `VERSION`（初版 `0.1.0`）をコピーして始める
- `VERSION` ファイルは 1 行 `X.Y.Z` のみ（末尾改行 1 個）

### G. リリース手順（ユーザー明示指示時のみ）

ユーザーが「リリース」「カーネル版を上げる」等の指示を出した場合のみ:

1. `kernel/KERNEL_VERSION` を新しい `X.Y.Z` に更新
2. `CHANGELOG.md` の `[Unreleased]` を `[X.Y.Z] - YYYY-MM-DD`（当日日付）に昇格し、直上に空の `[Unreleased]` を再設
3. 以下 3 箇所の版表記を `X.Y`（MAJOR.MINOR）に同期:
   - `README.md` L1 `# Fabriq ver{X.Y}`
   - `kernel/common.ps1` L2 `# Easy Kitting Batch - Common Function Library v{X.Y}.Z`（Z まで含める）
   - `kernel/main.ps1` L3 `# Fabriq ver{X.Y} - Manifeste du Surkitinisme -`
4. `pwsh ./dev/check_version.ps1` を実行して整合性確認
5. `git tag kernel-vX.Y.Z` は **Claude 側では実行しない**。コマンドを提示してユーザーにお願いする（モジュール単独リリースなら `git tag <module>-vX.Y.Z` を提示）

**ユーザーからの明示指示なしに `KERNEL_VERSION` / 版表記を勝手に昇格しないこと。** `[Unreleased]` への追記とモジュール `VERSION` の個別昇格は日常作業として進めてよい。

### H. 整合性チェック

- `pwsh ./dev/check_version.ps1` で `KERNEL_VERSION` と各ファイル版表記の整合を検証
- 非 0 終了した場合は **必ず** 版表記を揃えてからコミット
