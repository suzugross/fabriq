# ppkg_config (Standard)

**カテゴリ**: System
**メニュー名**: PPKG Install / PPKG Uninstall
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（cmdlet 例外有無のみで判定）
**サブスクリプト**: `ppkg_install_config.ps1`（適用）, `ppkg_uninstall_config.ps1`（削除）

## 目的
プロビジョニングパッケージ（.ppkg）の適用・削除を行う汎用モジュール。
`file/` ディレクトリに配置した PPKG ファイルを CSV 定義に従って一括処理し、
PackageName での冪等性チェックや Uninstall 時の物理ファイル削除リトライも備えます。
fabriq の標準モジュールでカバーしきれないレガシー設定（Provisioning Configuration Designer で
作成した独自設定）を一括導入する逃げ道としての位置付けです。

## 入力 (CSV)
`ppkg_list.csv`（Install / Uninstall 共通）:
- `Enabled`: 有効フラグ
- `PackageName`: パッケージ識別名（`Get-ProvisioningPackage` の PackageName と一致）
- `FileName`: `file/` 内の .ppkg ファイル名
- `Description`: 表示用
- `Segment`: Segment フィルタ

## 主要ステップ（Install）
1. CSV 読み込み + Enabled=1 抽出
2. `Install-ProvisioningPackage` cmdlet 存在確認 + `file/` ディレクトリ確認
3. ドライラン: 各エントリに [INSTALL] / [REINSTALL] / [NOT FOUND] マーカー表示
   （`Get-ProvisioningPackage` で PackageName 照合、ファイルサイズも表示）
4. 実行確認（AutoPilot は自動 Y）
5. ループ: `Install-ProvisioningPackage -QuietInstall -ForceInstall -PackagePath <ppkg>`
6. `New-BatchResult` で集計

## 主要ステップ（Uninstall）
1. CSV 読み込み + cmdlet 存在確認（`Get-ProvisioningPackage` / `Remove-ProvisioningPackage`）
2. ドライラン: [INSTALLED] / [NOT FOUND] 表示（PackageId も併記）
3. 実行確認
4. ループで再クエリ（dry-run と実行間で状態変化を考慮）→
   - Phase 1: `Remove-ProvisioningPackage -PackageId` で cmdlet 削除
   - Phase 2: `$pkg.PackagePath` の物理ファイルを 5 回まで 2 秒間隔リトライ削除
   - Phase 3: cmdlet 成功 / ファイル削除のみ成功 / 既に clean / 失敗 で結果分類
5. `New-BatchResult` で集計

## 注意点・運用メモ
- 管理者権限必須
- 適用後、設定の反映にはリブートまたは再ログオンが必要な場合あり
- `PackageName` は PPKG ビルド時に Provisioning Configuration Designer で
  指定した名前と完全一致させる必要あり（`Get-ProvisioningPackage` で確認可能）
- Install は `-ForceInstall` 付きで毎回再適用される（冪等チェックは表示のみ）
- Uninstall の物理ファイルロック対策（5 回リトライ）は Spooler 等の
  PPKG 参照プロセスが残っている状況を想定

## 検証
Post-Apply Verification は未実装。`-Verified` 引数は未渡しのため実行履歴の
Verified 列は空欄。cmdlet 例外の有無のみで成否判定するため、
PPKG が「適用されたが期待通りに動作していない」ケースは検出できない。
詳細検証が必要な場合は、PPKG が変更したレジストリ/ファイルを別モジュールで
事後確認する Profile 設計が推奨。
