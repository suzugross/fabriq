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
