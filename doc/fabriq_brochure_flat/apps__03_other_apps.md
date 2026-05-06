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
