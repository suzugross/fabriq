# sysprep_config (Standard)

**カテゴリ**: System
**メニュー名**: Sysprep Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 意図的に非対応（`/shutdown` モードで自プロセスが消えるため検証不可）
**サブスクリプト**: なし（CSV から `unattend.xml` / `SetupComplete.cmd` を動的生成）

## 目的
Sysprep の応答ファイル `unattend.xml` と初回起動スクリプト `SetupComplete.cmd` を 3 つの CSV
（`sysprep_list.csv` / `unattend_list.csv` / `setupcomplete_list.csv`）から動的生成・配置し、
最後に `sysprep.exe /generalize` を実行するモジュール。キッティングの「最終一段前」位置に配置し、
これ以降のイメージ封入工程の入口として機能する。第 1 確認（ファイル生成）と
第 2 確認（Sysprep 実行）の 2 段階確認を持ち、第 2 確認 N 選択時はファイル配置済み・未実行で
Status=Success を返して終了する。

## 入力 (CSV)
- `sysprep_list.csv`: `Enabled` / `SysprepExe`（フルパス）/ `Mode` (`oobe`/`audit`) /
  `Shutdown` (`shutdown`/`reboot`/`quit`) / `Description` / `Segment`
- `unattend_list.csv`: `Enabled` / `SettingName` (`ComputerName` / `CopyProfile` / `TestUserName` /
  `EnableAdministrator` / `AdminPassword` / `HideEULAPage` / `ProtectYourPC` /
  `HideWirelessSetupInOOBE` / `HideOnlineAccountScreens` / `HideOEMRegistrationScreen`) /
  `Value` / `Description` / `Segment`
- `setupcomplete_list.csv`: `Enabled` / `Order` / `ActionType` (`DeleteUser` / `CopyFile` /
  `Command`) / `Target` / `Destination` / `Description` / `Segment`
- `source/`: `CopyFile` アクションで使うファイル群を置くディレクトリ

すべて UTF-8 BOM 必須（`feedback_ps1_utf8_bom.md` の方針と整合）。

## 主要ステップ
1. 3 つの CSV を読み込み（sysprep は有効 1 行のみ、unattend はキー単位適用、setupcomplete は
   1 件以上）
2. 前提チェック: `sysprep.exe` 存在 / `source/` 存在 / `C:\Windows\Setup\Scripts\` 自動作成
3. 一覧表示（Sysprep 設定 / Unattend キー / SetupComplete アクション、`[NEW]`/`[OVERWRITE]`）
4. 第 1 確認: ファイル生成・配置（AutoPilot 自動 Y）
5. 配置:
   - 5-1. `source/` → `C:\Windows\Setup\Scripts\source\` ステージング (xcopy)
   - 5-2. `unattend.xml` 動的生成 → `C:\Windows\System32\Sysprep\unattend.xml`
   - 5-3. `SetupComplete.cmd` 動的生成 → `C:\Windows\Setup\Scripts\SetupComplete.cmd`
     （先頭でログを `SetupComplete.log` に追記）
6. 第 2 確認: Sysprep 実行（N 選択時はファイル配置済 + Status=Success で return）
7. `sysprep.exe /generalize /<mode> /<shutdown>` 実行
   （`shutdown`/`reboot` は PC 停止のため Step 7 後は到達せず、`quit` のみ ExitCode で判定）

## 注意点・運用メモ
- 管理者権限必須
- `<LocalAccount>` ブロック（TestUserName 指定時に生成）が unattend にあると Windows 仕様で
  アカウント関連 OOBE 画面が個別設定に関わらず自動スキップされる。OOBE を表示させたい場合は
  TestUserName 行を Enabled=0 にする
- `<AdministratorPassword>` はパスワード設定のみ。Administrator アカウント有効化は
  `setupcomplete_list.csv` の `net user Administrator /active:yes` を Enabled=1 にする
- 冪等性なし（実行のたびに上書き）。Sysprep 自体に同一イメージへの実行回数制限あり（既定 3 回）
- Default プロファイルを露出するための ATTRIB -H / 内部キャッシュ削除 / ATTRIB +H 一連の
  アクションが setupcomplete に標準同梱されている（INetCache / WebCache / LocalLow 削除）
- `taskbar_config` から自動コピーされる `LayoutModification.xml` が source/ 経由で
  Default User の Shell 配下へ展開されるという連携あり

## 検証
意図的に Post-Apply Verification 非対応（`project_verification_exclusions.md` 整合）:
`/shutdown` `/reboot` モードでは sysprep.exe 完了 = PC 停止のためスクリプト側に検証ステップを
走らせる余地がない。`/quit` モードでは検証可能だが通常運用では使わないため統一的に未実装。
`-Verified` を渡していないため履歴 Verified 列は常に空欄。

完了確認は手動で:
- `C:\Windows\System32\Sysprep\Panther\setuperr.log`
- `C:\Windows\Setup\Scripts\SetupComplete.log`
- 配置された `unattend.xml` / `SetupComplete.cmd`
