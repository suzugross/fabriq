# restore_point (Standard)

**カテゴリ**: System
**メニュー名**: Restore Point
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（事前冪等性チェックのみ。`-Verified` 未渡し）
**サブスクリプト**: なし（ローカルヘルパー `Test-RestoreRegistryValue` / `Get-ShadowStorageInfo`）

## 目的
Windows のシステムの保護（System Restore）まわりを CSV ベースで一括設定するモジュール。
保護有効化、24 時間制限解除、シャドウコピー容量設定、復元ポイント作成の 4 操作を 1 行 1 操作で
列挙する設計で、キッティング開始前のスナップショット取得などに利用する。クライアント OS
（Windows 10 / 11）専用で、サーバー OS では `Enable-ComputerRestore` 等が使えないため動作不可。

## 入力 (CSV)
`restore_point_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `SettingName`: 操作種別 (`enable_protection` / `remove_24h_limit` / `set_storage_size` / `create_restore_point`)
- `Drive`: 対象ドライブ（例 `C:\`、不要操作は空欄）
- `Description`: 説明（表示・復元ポイント名兼用）
- `Value`: 操作別値（容量 % / RestorePointType 等）
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 保護有効化 (C:\) → 24h 制限解除 → シャドウ容量 5% → fabriq キッティング前
ポイント作成 (MODIFY_SETTINGS) の 4 行が Enabled=1。

## 主要ステップ
1. `restore_point_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. `Test-AdminPrivilege` で管理者権限確認
3. ドライラン表示（操作別に冪等性チェック → `[SKIP]` / `[APPLY]` 表示、
   `vssadmin list shadowstorage` で現容量も併記）
4. 実行確認（AutoPilot 自動 Y）
5. 適用ループ (switch で operation 分岐):
   - `enable_protection`: `DisableSR=0` なら Skip、それ以外は `Enable-ComputerRestore -Drive`
   - `remove_24h_limit`: `SystemRestorePointCreationFrequency=0` なら Skip、それ以外は同値書き込み
   - `set_storage_size`: `vssadmin resize shadowstorage /maxsize=N%`（毎回実行、非冪等）
   - `create_restore_point`: `Checkpoint-Computer -Description -RestorePointType`（非冪等）
6. `New-BatchResult` 返却（`-Verified` は渡さない）

## 注意点・運用メモ
- 管理者権限必須、クライアント OS 限定
- CSV は論理依存順（保護有効化 → 制限解除 → 容量設定 → ポイント作成）の並びを推奨
- `set_storage_size` は vssadmin の制約で冪等チェックなし（毎回呼ぶ）
- `create_restore_point` は OS 側の 24 時間制限で実体としては Skip されるケースあり
  （事前に `remove_24h_limit` を実行する設計）
- `Test-RestoreRegistryValue` は `reg_hklm_config` の `Test-RegistryValueMatch` パターンを
  踏襲したローカルヘルパー、`Get-ShadowStorageInfo` は `ssid_config` の netsh パース流儀を踏襲
- 復元ポイントは Description がそのまま表示名になるため CSV の Description 欄は意味のある日本語可

## 検証
operation 別に挙動が異なる:
- `enable_protection` / `remove_24h_limit`: 適用前にレジストリ読み返しで一致なら Skip と判定。
  これが事実上の事前検証だが、Step 5.5 として独立した読み返しは実装していない
- `set_storage_size`: vssadmin 戻り値（`$LASTEXITCODE`）のみで判定。容量設定の最終値読み返し未実装
- `create_restore_point`: `Checkpoint-Computer` 例外なしを成功として扱う

`New-BatchResult` に `-Verified` を渡していないため履歴の Verified 列は空欄。
実効性は手動で `vssadmin list shadowstorage` / `Get-ComputerRestorePoint` 等で確認する運用。
