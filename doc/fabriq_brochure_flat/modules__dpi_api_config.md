# dpi_api_config (Standard)

**カテゴリ**: Display
**メニュー名**: DPI Scaling Config (Live)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（適用前の冪等性チェックには `GetCurrentDpi` を使用するが、適用後の読み返し検証は省略）
**サブスクリプト**: なし

## 目的
ディスプレイの **DPI スケーリング（拡大率）を Windows API 経由で即時反映** させるモジュールです。レジストリ書き込み + サインアウト方式（`resolution_api_config` 系統）と異なり、`NativeDpiHelper::SetDpi` を呼び出すライブ方式のため、再起動／サインアウトなしで即座に変更が反映されるのが特徴です。「Live」というメニュー名はこのライブ反映を表しています。複数モニター環境では `MonitorIndex`（0=プライマリ）で個別に指定できます。

## 入力 (CSV)
`dpi_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `MonitorIndex` … 0=プライマリ、1=セカンダリ、…
- `ScalePercent` … 拡大率（100 / 125 / 150 / 175 / 200 等、Windows がサポートする刻み）
- `Description` / `Segment`

## 主要ステップ
1. `dpi_list.csv` 読み込み（Enabled=1）
2. 対象モニターと拡大率を表示
3. （冪等性チェック）`NativeDpiHelper::GetCurrentDpi` で現在値読み出し、目標と一致なら Skip
4. 実行確認（AutoPilot 時は自動 Y）
5. Windows API 経由で DPI を変更（即時反映）

## 注意点・運用メモ
- ノート PC の高 DPI 表示に合わせて 150% を設定するなど、現場では頻出のオペレーション
- レジストリ書き込み方式（`resolution_api_config` の dpi_reg 系）と違い、再起動／サインアウト不要
- 環境変数は使用しない
- Windows のサポート刻みから外れた `ScalePercent` を指定すると API 側で丸め or 失敗する可能性あり

## 検証
Post-Apply Verification は **未実装**。`NativeDpiHelper::GetCurrentDpi` は **適用前の冪等性チェック**（既に目標値ならスキップ）にのみ使用し、適用後の読み返し検証は行いません。理由は、API 反映直後と OS 内部のメトリック更新の間にラグがあり、即座の読み返しは false negative を出しやすいためです。`-Verified` は未渡しで Verified 列は空欄。実反映は次回ログオン後の見た目で確認する運用です。
