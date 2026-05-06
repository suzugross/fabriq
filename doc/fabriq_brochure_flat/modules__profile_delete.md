# profile_delete (Standard)

**カテゴリ**: User Management
**メニュー名**: Delete User Profiles
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（実装推奨メモは Guide にあり、現状未実装）
**サブスクリプト**: `profile_delete.ps1`（メイン処理 1 本のみ）

## 目的
指定ユーザーのプロファイル（`%SystemDrive%\Users\<UserName>`）を削除するモジュール。
WMI（`Win32_UserProfile`）を使用して Windows のプロファイル登録を正しく解除し、
WMI レコードが存在しない孤児フォルダや WMI 削除後にフォルダだけが残ったケースに
備えて `Remove-Item -Recurse -Force` の物理削除フォールバックを持ちます。
キッティング後に testuser や defaultuser0 を一掃する用途が代表的。

## 入力 (CSV)
`profile_list.csv`:
- `Enabled`: 有効フラグ（1=削除対象, 0=スキップ）
- `UserName`: ユーザー名（`C:\Users\<UserName>` のフォルダ名と一致）
- `Description`: 表示用
- `Segment`: Segment フィルタ

## 主要ステップ
1. 管理者権限チェック（`Test-AdminPrivilege`）
2. `profile_list.csv` 読み込み（Enabled=1 のみ、UserName 必須）
3. `%SystemDrive%\Users` 存在確認（D: ドライブインストール等にも追従）
4. ドライラン: 各エントリの実フォルダ存在を [APPLY] / [SKIP] と表示
5. 実行確認（AutoPilot は自動 Y）
6. 削除ループ:
   - フォルダ不在は Skip
   - **Stage 1**: `Get-CimInstance Win32_UserProfile | Where { LocalPath -like "*\$userName" }` で
     WMI レコードを取得し `Remove-CimInstance` で正規登録解除
   - **Stage 2**: フォルダがまだ残っていれば `Remove-Item -Path -Recurse -Force` で物理削除
7. `New-BatchResult` で集計

## 注意点・運用メモ
- 管理者権限必須（`Win32_UserProfile.Remove-CimInstance` 実行）
- 現在ログオン中のユーザープロファイルは削除不可（OS が拒否、Error 計上）
- WMI 削除のみではフォルダが残るケースがあり、Stage 2 のフォールバックは必須
- 削除されたユーザーの SID / プロファイル情報が Windows から登録解除される（レジストリも含む）
- パス検証（destructive_path_safety）: `Join-Path $usersBase $userName` の結果が
  `$usersBase` 配下であることを暗黙に前提（CSV の UserName に `..\` 等の
  path traversal が混入した場合の防御は弱い、運用 CSV のレビュー前提）

## 検証
Post-Apply Verification は未実装。Guide には実装推奨メモがあり、
- `Test-Path "$env:SystemDrive\Users\<UserName>"` が `$false`
- `Get-CimInstance Win32_UserProfile` で該当 LocalPath 不在

の 2 条件で `-Verified $true` 返却が可能との設計指針が記載されているが、現状の
`New-BatchResult` には `-Verified` を渡していない。実行履歴の Verified 列は空欄。
