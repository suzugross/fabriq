# power_config (Standard)

**カテゴリ**: Power
**メニュー名**: Power Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Win32 API による全設定値読み返し、タイムアウトは ±1 分許容）
**サブスクリプト**: `power_config.ps1`（メイン処理 1 本のみ。P/Invoke で powrprof.dll を直接呼び出し）

## 目的
Windows の電源プラン（PowerPlan）と電源モードオーバーレイ（PowerMode、Win11 のスライダー）、
そしてディスプレイ OFF / スリープ / 休止 / HDD オフ / 電源ボタン動作 / カバー閉じ動作 /
プロセッサ最小・最大状態を一括設定する集約モジュール。
ボタン/カバー設定はコントロールパネル UI と同じ Win32 API（`PowerWriteACValueIndex` 等）を
P/Invoke で直接呼び出すため、`powercfg.exe` が機能しない HP 等 OEM PC でも正しく動作します。

## 入力 (CSV)
`power_list.csv`（プロファイル単位の行ベース、Enabled=1 を最初に発見した行を自動採用）:
- `Enabled`, `ProfileName`, `Description`
- `PowerPlan`: BALANCED / HIGH_PERFORMANCE / POWER_SAVER
- `PowerMode`: BEST_PERFORMANCE / BALANCED / BEST_EFFICIENCY（オーバーレイ、空欄/`-` で変更なし）
- `Display_TurnOff_AC` / `_Battery`: ディスプレイ OFF までの分（0=無効）
- `Sleep_After_AC` / `_Battery`, `Hibernate_After_AC` / `_Battery`
- `HardDisk_TurnOff_AC` / `_Battery`
- `PowerButton_AC` / `_Battery`, `SleepButton_AC` / `_Battery`, `LidClose_AC` / `_Battery`:
  NOTHING / SLEEP / HIBERNATE / SHUTDOWN
- `Processor_MinState_AC` / `_Battery`, `Processor_MaxState_AC` / `_Battery`: %
- `Segment`

## 主要ステップ
1. CSV 読み込み + Enabled=1 の最初の行を自動選択（無ければ対話メニュー）
2. 設定内容表示 + 確認
3. PowerPlan 切替（`powercfg /S <GUID>`）+ 冪等チェック（GETACTIVESCHEME と一致なら Skip）
4. PowerMode オーバーレイ（`PowerSetUserConfiguredACPowerMode` / `DC` API）
5. 各タイムアウト設定: `PowerReadACValueIndex` / `DC` で現在値読み返し → 一致なら Skip、
   不一致なら `powercfg /CHANGE` または `/SETACVALUEINDEX` で適用
6. ボタン/カバー設定: Win32 API `PowerWriteACValueIndex` / `DC` で直接書き込み
   （powercfg がサイレント失敗する OEM 制限を回避）
7. プロセッサ設定: 同じく Win32 API で AC/DC 別々に書き込み
8. `PowerSetActiveScheme` で確定反映
9. Step 5.5: Post-Apply Verification（後述）
10. `New-ModuleResult -Verified` で結果返却

## 注意点・運用メモ
- 管理者権限必須（`#Requires -RunAsAdministrator`）
- 冪等性チェックは `powercfg /QUERY` のテキスト解析ではなく Win32 API ReadValueIndex を使用
  → OS locale 非依存（日本語 Windows でも正しく動作）
- HP 製 PC: コントロールパネル「カバーを閉じたときの動作」表示が Win32 API 経由設定後も
  更新されない既知問題あり。実際の挙動とレジストリ値は正しく設定済み（HP 固有サービスの表示問題）
- Hibernate のタイムアウトは Windows 内部調整で ±1 分の誤差が出るため tolerance あり

## 検証
Post-Apply Verification を実装。Power Plan GUID、Power Mode（AC/DC 各々）、
全タイムアウト 8 種、ボタン/カバー 6 種、プロセッサ状態 4 種を Win32 API で
読み返し、CSV 期待値と照合。タイムアウト系は ±1 分 tolerance、その他は完全一致。
全件 PASS で `-Verified $true`、1 件でも不一致で `$false`、適用対象なしで `$null` を返却。
