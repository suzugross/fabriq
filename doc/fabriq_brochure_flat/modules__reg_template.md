# reg_template (Extended)

**カテゴリ**: System
**メニュー名**: Registry Backup (Template) / Registry Import (Template)（2 メニュー）
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（reg.exe ExitCode 判定のみ、`.reg` 内容や復元後の値読み返しなし）
**サブスクリプト**: `reg_backup.ps1`（エクスポート）+ `reg_import.ps1`（インポート）。`backup/` ディレクトリにタイムスタンプ付き `.reg` を蓄積

## 目的
任意のレジストリキーを `.reg` ファイルにバックアップ／復元するための**テンプレートモジュール**（モジュール名に "Template" を冠する通り）。`desktop_icon_config` は本モジュールを特定キー専用に固定したサンプルと位置付けられ、本テンプレートは「ディレクトリをコピー → reg_list.csv を編集して任意キー対応モジュールを量産」する出発点。

## 入力 (CSV)
`reg_list.csv`:
- **Enabled**: 有効フラグ
- **RegistryPath**: フルパス（`HKEY_LOCAL_MACHINE\...` 形式）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ（Backup）
1. CSV 読み込み
2. `backup/` 必要に応じて作成
3. 各レジストリパスの存在確認 + リスト表示（`[OK]`/`[--]` マーカー）
4. `Confirm-ModuleExecution`
5. 各キーに対し `reg.exe export` を実行、ファイル名は `yyyyMMdd_HHmmss_<sanitized_path>.reg`（パス内の `\/:*?"<>|` を `_` に置換、長すぎる場合は 80 文字に切り詰め）
6. `New-BatchResult` 集計

## 主要ステップ（Import）
1. CSV 読み込み
2. `backup/` 存在確認
3. 各 RegistryPath に対し `*_<sanitized_path>.reg` パターンでマッチする最新ファイルを `Sort-Object LastWriteTime -Descending` で特定
4. 表示（`[Ready]`/`[Missing]`、複数バックアップあれば「N backups available, using latest」表示）
5. 全件 Missing なら Error 終了
6. Warning 表示後 `Confirm-ModuleExecution`
7. 各 `.reg` を `reg.exe import`
8. `New-BatchResult` 集計（成功 1 件以上で「サインアウト/再起動が必要な場合あり」警告）

## 注意点・運用メモ
- 同一キーの過去バックアップが `backup/` 内に蓄積される（タイムスタンプで自然に世代管理）
- インポートは常に最新ファイルを使用、過去版に戻したい場合は手動でファイル名指定が必要（GUI 未提供）
- 管理者権限必須（HKLM 操作のため）
- ファイル名サニタイズの 80 文字制限により、極端に長いキーパスは衝突リスクあり

## 検証
未実装。両スクリプトとも `reg.exe` の ExitCode のみで成否判定。`-Verified` 未渡しで Verified 列は空欄。
