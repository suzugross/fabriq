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
