# display_config (Extended)

**カテゴリ**: Display
**メニュー名**: Display Resolution Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（適用反映が再起動後のため即時読み返し検証は無意味、ただし冪等性チェックは実装あり）
**サブスクリプト**: なし

## 目的
ディスプレイ解像度をレジストリ `HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration\<HardwareID>` 配下の `PrimSurfSize.cx/cy` 書き換えで設定するモジュール。
standard 側の `resolution_api_config`（DisplayConfig API による即時反映、単一ディスプレイのみ）と棲み分けて、こちらは**複数モニター・HardwareID 指定・再起動後反映**を担う extended 版。

## 入力 (CSV)
`display_list.csv`:
- **Enabled**: 有効フラグ
- **HardwareID**: ディスプレイ ID（`AUTO`=自動検出 / 具体的 ID は前方一致 / 空欄=対話選択）
- **Width** / **Height**: ピクセル
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み + Width/Height の正数検証
3. 対象ディスプレイ解決（AUTO=単一なら自動 / 複数 or 0 件は Interactive にフォールバック / 具体的 ID は前方一致）
4. ドライラン表示（`[Current]` / `[Change]` / `[Interactive]` マーカー、現在解像度併記）
5. `Confirm-ModuleExecution`
6. 各キーの最初のサブキーに対し冪等性チェック → `PrimSurfSize.cx/cy` を DWORD で書き込み
7. `New-BatchResult` 集計（成功 1 件以上で「再起動が必要」警告）

## 注意点・運用メモ
- 反映には**再起動が必須**
- AutoPilot 運用時は HardwareID を空欄/AUTO（複数モニター環境）にすると Interactive にフォールバックしてブロックされるため、具体的 ID 指定推奨
- 同じ Description で複数キーがマッチした場合はマッチした全キーに書き込む

## 検証
冪等性チェック（適用前の `PrimSurfSize.cx/cy` 読み出し）は実装ありだが、再起動が必要な性質上 Post-Apply Verification は省略。`-Verified` 未渡しで履歴の Verified 列は空欄。
