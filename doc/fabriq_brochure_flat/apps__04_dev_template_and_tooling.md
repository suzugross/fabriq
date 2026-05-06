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
