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
