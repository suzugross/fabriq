# startup_command_config (Standard)

**カテゴリ**: Scripts
**メニュー名**: Startup Command Deploy
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（将来の新規ユーザー初回ログオンで効果が出るため、`-Verified` 未渡し）
**サブスクリプト**: なし（CSV から `apply_startup_commands.ps1` を動的生成）

## 目的
新規ユーザーの初回サインイン時に一度だけ実行されるコマンド列を、CSV から動的生成して
Default User の Startup フォルダに仕掛けるモジュール。`reg_hkcu_config` / `spi_config` と
共通の Startup ランチャー (`fabriq_user_setup.ps1`) と Startup トリガー
(`FabriqUserSetup.cmd`) を共有し、すべての `apply_*.ps1` が初回ログオン時に同時に走り、
完了後に Explorer 再起動 + Startup トリガー自己削除という belt-and-suspenders 機構の一翼を担う。

## 入力 (CSV)
`startup_command_list.csv`
- `Enabled`: 1=有効 / 0=スキップ
- `Order`: 実行順（昇順数値ソート）
- `Description`: 表示用説明
- `Command`: 実行コマンド文字列（`cmd.exe /c` 経由）
- `Segment`: Segment フィルタ（任意）

例: タイムゾーン設定、フォルダ作成、追加 PowerShell スクリプト起動など。

## 主要ステップ
1. `startup_command_list.csv` を読み込み（Enabled=1）→ Order 昇順ソート
2. ドライラン表示（生成予定ファイル一覧）
3. 実行確認（AutoPilot 自動 Y）
4. CSV から `apply_startup_commands.ps1` を動的生成 → `C:\ProgramData\fabriq\` に配置
5. `Deploy-FabriqUserSetupLauncher` で `fabriq_user_setup.ps1`（共通ランチャー）配置
6. `Deploy-FabriqStartupTrigger` で `FabriqUserSetup.cmd` を Default User Startup に配置
7. `New-BatchResult` 返却（`-Verified` 未渡し）

新規ユーザー初回ログオン時の挙動: Startup → .cmd → フラグチェック（無し）→ PowerShell
起動 → ランチャー → `apply_*.ps1` 群（hkcu / spi / startup_commands）順次実行 →
Explorer 再起動 → `%LOCALAPPDATA%\fabriq\user_setup_done.flag` 作成 → .cmd 自己削除。

## 注意点・運用メモ
- 管理者権限必須（Default User Startup / ProgramData 書き込み）
- コマンドは新規ユーザーコンテキスト（非昇格）で実行されるため、HKLM 書き込みなどは不可
- 各コマンドは独立実行（1 つの失敗が他をブロックしない）
- 実行ログは `%LOCALAPPDATA%\fabriq\startup_commands.log`
- 再実行は冪等（最新 CSV 内容で `apply_*.ps1` を毎回上書き）
- 2 回目以降のログオンは .cmd 削除済 or フラグ検出で即終了するため、軽量
- `reg_hkcu_config` / `spi_config` の Active Setup / Startup Batch と機構を共有しているため、
  他 2 モジュールを使わなくても本モジュール単独で Startup Batch インフラを敷ける

## 検証
本モジュールに Post-Apply Verification は未実装。配備したトリガーが効くのは将来の
新規ユーザー初回ログオン時であり、キッティング時に読み返しても実効性は判定不能。
そのため `New-ModuleResult` / `New-BatchResult` に `-Verified` を渡しておらず、
履歴の Verified 列は空欄。実効性確認は実際に新規ユーザーアカウントを作成して
ログオンし、`%LOCALAPPDATA%\fabriq\startup_commands.log` を確認する手動運用。
