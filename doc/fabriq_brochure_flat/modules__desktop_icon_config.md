# desktop_icon_config (Extended)

**カテゴリ**: Desktop
**メニュー名**: Desktop Icon Backup / Desktop Icon Restore（2 メニュー）
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（Backup は ExitCode + ファイルサイズで成否、Restore はサインアウト後反映のため即時検証無意味）
**サブスクリプト**: `desktop_icon_backup.ps1`（エクスポート役）+ `desktop_icon_restore.ps1`（インポート役）。`backup/` ディレクトリにタイムスタンプ付き `.reg` を蓄積

## 目的
デスクトップアイコンの配置情報を保持するレジストリキー `HKCU\Software\Microsoft\Windows\Shell\Bags\1\Desktop` を `.reg` ファイルにバックアップ／復元するペアモジュール。
SYSTEM 起動セッションでも動作するように `Resolve-HkcuRoot` で対象ハイブを解決し、エクスポート時は `HKEY_USERS\<SID>` を `HKEY_CURRENT_USER` に正規化してポータブルな `.reg` を生成、インポート時は逆方向の書き換えを TEMP に作って `reg import` する仕組み。

## 入力 (CSV)
なし（対象レジストリパスはハードコード）。

## 主要ステップ（Backup）
1. `Resolve-HkcuRoot` で書き込み先ハイブ（HKCU or HKU\<SID>）を決定
2. 対象キーの存在確認 + 状態表示
3. `Confirm-ModuleExecution` で実行確認
4. `backup/` を必要なら新規作成
5. `reg.exe export` でタイムスタンプ付きファイル名 (`DesktopIcons_yyyyMMdd_HHmmss.reg`) に出力
6. HKCU リダイレクト時は出力 `.reg` 内の `HKEY_USERS\<SID>` を `HKEY_CURRENT_USER` に置換して可搬化

## 主要ステップ（Restore）
1. `backup/` から `DesktopIcons_*.reg` を新しい順に列挙（最大 5 件まで一覧表示）
2. 最新ファイルをターゲットとして表示
3. `Confirm-ModuleExecution` で実行確認
4. HKCU リダイレクト時は `.reg` の `HKEY_CURRENT_USER` を `HKEY_USERS\<SID>` に置換した一時ファイルを TEMP に作成
5. `reg.exe import` で書き戻し
6. 一時ファイルクリーンアップ（`finally`）

## 注意点・運用メモ
- バックアップ前にアイコンを希望の配置に並べておく必要あり
- 復元の反映にはサインアウトまたは再起動が必要
- 管理者権限必須（HKU 操作のため）
- 過去のバックアップは `backup/` 内に蓄積され続けるため、不要なら手動削除

## 検証
両スクリプトとも `reg.exe` の ExitCode 判定のみ。`.reg` の中身検証や、Restore 後のキー値読み返しは行わない。`-Verified` 未渡しで履歴の Verified 列は空欄。
