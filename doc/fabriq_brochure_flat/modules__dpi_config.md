# dpi_config (Extended)

**カテゴリ**: Display
**メニュー名**: Display DPI Scaling Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（書き込みごとに HKCU/HIVE 値の冪等性チェック実装あり、ただし -Verified 未渡し）
**サブスクリプト**: なし（C# `DpiScaleResolver` クラスを `Add-Type` で動的コンパイル）

## 目的
モニターごとの DPI スケーリング（拡大率 100/125/150/175/200%）を、HKCU と Default プロファイル ntuser.dat（HKU\Hive）両方に同時書き込みするモジュール。
standard の `dpi_api_config`（即時反映・単一ディスプレイのみ）と異なり、こちらは**複数モニター・HardwareID 指定・新規ユーザー継承**まで対応。
DisplayConfig API（`DisplayConfigGetDeviceInfo` 等）を P/Invoke で叩く C# クラス `DpiScaleResolver` を内包し、モニターごとの推奨 DPI（recommended）からの相対値で `DpiValue` を算出する点が技術的特徴。

## 入力 (CSV)
`dpi_list.csv`:
- **Enabled**: 有効フラグ
- **HardwareID**: `AUTO` / 具体的 ID（前方一致） / 空欄（Interactive Display）
- **ScalePercent**: 100/125/150/175/200 のいずれか / 空欄や 0（Interactive Scale）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` + `Resolve-HkcuRoot` で HKCU 解決（昇格セッション対応）
2. C# `DpiScaleResolver` を `Add-Type` で読み込み、各モニターの推奨 DPI マップを構築
3. CSV 読み込み + ScalePercent 値の検証（サポート値以外はスキップ、空欄は Interactive Scale 扱い）
4. 対象解決（PerMonitorSettings 検索 → なければ GraphicsDrivers\Configuration へフォールバック → AUTO 複数や空欄は Interactive へ）
5. ドライラン表示（`[Current]`/`[Change]`/`[Interactive]` 色分け、推奨値 Recommended 表示）
6. `Confirm-ModuleExecution`
7. `reg load HKU\Hive C:\Users\Default\ntuser.dat` で Default ハイブをロード
8. 書き込みループ（ScalePercent → DpiValue 換算 → HKCU と HIVE 各々で冪等性チェック → 書き込み）
9. `reg unload`（失敗時は GC 強制 + sleep してリトライ）
10. `New-BatchResult` 集計

## 注意点・運用メモ
- 反映にはサインアウト/再起動が必要
- Default プロファイル書き込みのため、本モジュールは Profile 全体（hostname → 各種設定 → 新規ユーザー作成）の一環として使う想定
- AutoPilot で完全自動にしたい場合は HardwareID と ScalePercent の両方を CSV で具体指定すること（Interactive にフォールバックすると入力待ちでブロック）
- 推奨 DPI 取得失敗時はフォールバック値 150% を使用

## 検証
HKCU と HIVE 双方で書き込み前に現在 `DpiValue` を取得して冪等判定（一致時 Skip）はあるが、Post-Apply Verification は未実装。`-Verified` 未渡しで履歴 Verified 列は空欄。手動検証は `Get-ItemProperty 'HKCU:\Control Panel\Desktop\PerMonitorSettings\<key>' -Name DpiValue`。
