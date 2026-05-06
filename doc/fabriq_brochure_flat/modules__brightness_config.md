# brightness_config (Standard)

**カテゴリ**: Display
**メニュー名**: Brightness Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`WmiMonitorBrightness.CurrentBrightness` を読み返し）
**サブスクリプト**: なし

## 目的
ノート PC のキッティング工程で「内蔵ディスプレイの初期輝度を一定値（例: 80%）に揃える」という標準作業を CSV 1 行で表現するモジュールです。WMI クラス `root\wmi\WmiMonitorBrightnessMethods` を介して即時反映するため、再起動を要しません。外付けモニターのみで内蔵ディスプレイを持たないデスクトップ PC では WMI が応答しない仕様のため、本モジュールは **Error ではなく Skipped で安全終了** する設計になっており、共通プロファイルでデスクトップ／ノートを混載しても運用しやすいようになっています。

## 入力 (CSV)
`brightness_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Brightness` … 輝度（0〜100）
- `Description` … 表示用ラベル
- `Segment` … Segment フィルタ

## 主要ステップ
1. CSV 読み込み（`Import-ModuleCsv -FilterEnabled`）
2. WMI 輝度制御の対応可否確認（未対応なら Skipped で終了）
3. 現在の輝度と変更先を表示（dry-run）
4. 実行確認（AutoPilot 時は自動 Y）
5. WMI メソッドで輝度変更
5.5 **Post-Apply Verification**: `WmiMonitorBrightness.CurrentBrightness` を読み返し、CSV 目標値と一致するか検証
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- ノート PC など内蔵ディスプレイを持つデバイスでのみ動作
- 外付けモニターのみのデスクトップ PC は自動 Skipped で安全終了（Error 扱いにしない）
- 設定値は OS の電源プラン UI における「100% / 80% / 50%」スライダ位置に直接対応
- 環境変数は使用しない

## 検証
Post-Apply Verification は **実装あり**。WMI の `CurrentBrightness` を再取得して CSV 目標値との完全一致を判定し、一致なら Verified=true、不一致なら false で `New-BatchResult` に `-Verified` フラグを返却します。WMI 経由の即時反映なので、書き込み直後の読み返しが事実上信頼できる検証になっています。
