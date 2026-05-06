# partition_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: Partition Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（パーティション存在 + FileSystem + サイズ ±5% 許容）
**サブスクリプト**: `partition_config.ps1`（メイン処理 1 本のみ）

## 目的
PowerShell Storage コマンドレット（`Resize-Partition` / `New-Partition` / `Format-Volume`）を
使用して既存パーティションの縮小と新規パーティション作成を行うモジュール。
複数パーティション分割（C → D → E）に対応し、CD-ROM 等が削除予定のドライブレターを
占有している場合は自動で別レターに退避（Z..Q から空きを探す）する安全機構を持ちます。
hostlist 連動ではなくモジュール CSV 駆動。

## 入力 (CSV)
`partition_list.csv`:
- `Enabled`: 有効フラグ
- `DiskNumber`: 対象ディスク番号（通常 0）
- `SourceDriveLetter`: 縮小対象（C 等）
- `SourceSizeMB`: 縮小後サイズ（MB）
- `NewDriveLetter`: 新規パーティションのドライブレター
- `NewSizeMB`: 新規サイズ（0=残り全領域、最終行のみ可）
- `FileSystem`: NTFS / ReFS
- `VolumeLabel`: ボリュームラベル
- `Description`, `Segment`

## 主要ステップ
1. CSV 読み込み（Enabled=1）
2. バリデーション:
   - `NewSizeMB=0`（残り全領域）は最大 1 行 + 必ず最終行
   - `NewDriveLetter` 重複チェック
3. ドライラン表示（[APPLY] / [SKIP] / [RELOCATE]）
4. 実行確認（AutoPilot は自動 Y）
5. **Phase A 縮小**: Source ごとに重複排除し、現サイズ ≦ 目標なら Skip、
   それ以外は `Get-PartitionSupportedSize` で SizeMin チェック後 `Resize-Partition`
6. **Phase B 新規作成**: ドライブレター衝突を検知 → CD-ROM 等は `Move-ConflictingDriveLetter` で退避、
   既存パーティションなら Skip、`New-Partition` + `Format-Volume`
7. Step 5.5: Verification（Source サイズ + 各 New パーティションの存在/FS/サイズ）
8. `New-BatchResult -Verified` で集計返却

## 注意点・運用メモ
- 管理者権限必須（Storage コマンドレット全般）
- 縮小対象ドライブが他プロセスにロックされていると失敗
- CD-ROM が D: を占有している場合の自動退避は Z, Y, X, W, V, U, T, S, R, Q から
  空きを順次選択（10 候補すべて埋まっていれば Error）
- Resize の `SizeMin` 制約は Windows ファイルシステムの未移動データ依存で、
  デフラグ未実施の場合は予想より大きい数字になることがある
- 冪等性は NewDriveLetter ベースで判定（Source の縮小は ≦ チェックで保護）

## 検証
Post-Apply Verification を実装。各 Source パーティションの実サイズが目標 ±5% 以内、
各 New パーティションの存在 / FileSystem 一致 / サイズ ±5% 以内（NewSizeMB>0 のみ）を
読み返し、全合致で `-Verified $true`、不一致 1 件以上で `$false` を返却。
ディスク操作直後の OS 内部状態安定化を考慮した tolerance 設計。
