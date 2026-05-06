# volume_config (Standard)

**カテゴリ**: System
**メニュー名**: Volume Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Core Audio API で読み返し）
**サブスクリプト**: なし（C# / COM 経由で `IAudioEndpointVolume` を `Add-Type` で同居）

## 目的
PC 本体のマスター音量とミュート状態を Windows Core Audio API（MMDeviceAPI /
IAudioEndpointVolume）で設定するモジュール。タスクバーの音量スライダーに即時反映され、
再起動不要。受け取り検査時の「全 PC が同じデフォルト音量で出荷される」ことや、
キッティング作業中の意図せぬ大音量再生を抑止する用途で使う。

## 入力 (CSV)
`volume_list.csv`
- `Enabled`: 1=適用 / 0=スキップ
- `Volume`: 0〜100 の整数（%）
- `Mute`: `on` / `off` / 空欄（変更しない）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 1 行 `Volume=50, Mute=off`（マスター音量 50%）。
有効行は最初の 1 行のみ採用される設計（複数行を意味的に重ねない）。

## 主要ステップ
1. `volume_list.csv` 読み込み（最初の Enabled=1 行のみ）
2. オーディオデバイス存在確認（VM でオーディオ無効なら安全終了）
3. `IAudioEndpointVolume::GetMasterVolumeLevelScalar` / `GetMute` で現在値取得・表示
4. 冪等性チェック（既に目標値一致なら Skip）
5. 実行確認（AutoPilot 自動 Y）
6. `SetMasterVolumeLevelScalar` / `SetMute` で即時設定
7. Step 5.5: `GetMasterVolumeLevelScalar` / `GetMute` で読み返し検証 → CSV 期待値と一致確認 →
   `New-ModuleResult -Verified $verified` 返却

## 注意点・運用メモ
- オーディオデバイスが存在しない VM 環境では COM 取得時にエラーになるが安全終了する設計
- デフォルトオーディオデバイスのみ対象（複数デバイスが刺さっている場合は OS 既定デバイス）
- 音量値は scalar (0.0-1.0)、CSV の % 表記は内部で /100 して API に渡す
- Mute=空欄 のときはミュート状態は変更しない（音量のみ変える運用）
- 即時反映なので、キッティング BGM 等を流しているとその場で音量が変わる副作用あり

## 検証
Step 5.5 で `IAudioEndpointVolume::GetMasterVolumeLevelScalar` を再呼び出しして
小数誤差を吸収しつつ ±1% 以内で一致を判定（scalar↔% の往復で丸め誤差が出るため）。
`GetMute` で BOOL を再取得して CSV 値と比較。両方一致で `-Verified=$true`。
即時反映 API なので Verification がそのまま適用結果を反映するパターン。
