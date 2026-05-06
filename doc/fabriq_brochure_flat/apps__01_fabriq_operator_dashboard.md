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
