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
