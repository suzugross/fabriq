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

## バージョン管理ルール（CHANGELOG 駆動）

本プロジェクトは `VERSION` ファイル（ルート直下）を**単一の真実のソース**とし、
[Semantic Versioning 2.0.0](https://semver.org/lang/ja/) に準拠します。
変更履歴は [`CHANGELOG.md`](CHANGELOG.md) で [Keep a Changelog 1.1.0](https://keepachangelog.com/ja/1.1.0/) 形式に従います。

### A. コード変更時の義務（毎回）

`kernel/` / `modules/` / `apps/` / `commands/` / `profiles/` / `dev/template/` 配下のコードまたは CSV スキーマを変更した場合、**同じコミット内で必ず** 以下を実施してください:

1. `CHANGELOG.md` の `[Unreleased]` セクションに該当項目を追記する
2. カテゴリ（`Added` / `Changed` / `Deprecated` / `Removed` / `Fixed` / `Security`）を正しく選ぶ
3. 行頭にモジュール名またはコンポーネント名のプレフィックスを付ける
   - 例: `- hostname_config: Post-Apply Verification を追加`
   - 例: `- kernel/common.ps1: Import-ModuleCsv に -Segment パラメータ追加`
4. 1行80文字を目安に簡潔に。複数行が必要な場合はサブ項目で展開する

ドキュメント（README, Guide.txt, コメントのみ）の修正は追記不要ですが、追記しても構いません。

### B. バージョン番号の影響判定（昇格時に参照）

| 影響 | インクリメント | 例 |
|---|---|---|
| **MAJOR** (X+1.0.0) | 破壊的変更 | `kernel/common.ps1` 公開関数の削除・シグネチャ変更 / `module.csv` / プロファイル CSV の必須列削除・改名 / `New-ModuleResult` / `New-BatchResult` の返却契約変更 / 環境変数規約（`SELECTED_*` 等）の破壊的変更 |
| **MINOR** (0.Y+1.0) | 後方互換な追加 | 新規モジュール追加 / `common.ps1` への公開関数追加 / `module.csv` / プロファイル CSV への任意列追加 / `preset.csv` 対応拡張 |
| **PATCH** (0.0.Z+1) | バグ修正のみ | 既存挙動の修正（外部仕様不変） / ログ文言・コメント・Guide.txt 修正 / エビデンス出力のみの軽微変更 |

判定に迷ったら**大きい側に倒す**こと。

### C. リリース手順（ユーザーの明示指示があった時のみ実施）

ユーザーが「リリースする」「版を上げる」等の指示を出した場合のみ、以下を Claude Code 側で実施:

1. `VERSION` を新しい `X.Y.Z` に更新
2. `CHANGELOG.md` の `[Unreleased]` を `[X.Y.Z] - YYYY-MM-DD`（当日日付）に昇格し、直上に空の `[Unreleased]` を再設
3. 以下3箇所の版表記を `X.Y`（MAJOR.MINOR のみ）に同期:
   - `README.md` L1 `# Fabriq ver{X.Y}`
   - `kernel/common.ps1` L2 `# Easy Kitting Batch - Common Function Library v{X.Y}.Z`（Z まで含める）
   - `kernel/main.ps1` L3 `# Fabriq ver{X.Y} - Manifeste du Surkitinisme -`
4. `pwsh ./dev/check_version.ps1` を実行して整合性確認
5. `git tag vX.Y.Z` は **Claude Code 側では実行しない**。コマンドを提示してユーザーにお願いする

**ユーザーからの明示指示なしに `VERSION` / 版番号表記を勝手に変更しないこと。** `[Unreleased]` への追記のみが日常の作業。

### D. 整合性チェック

- `dev/check_version.ps1` を手動実行することで、`VERSION` と各ファイルの版表記が揃っているか検証できる
- 非0終了した場合は **必ず** 版表記を揃えてからコミットする
