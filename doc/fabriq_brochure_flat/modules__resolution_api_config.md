# resolution_api_config (Standard)

**カテゴリ**: Display
**メニュー名**: Resolution Config (Live)
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（事前判定のみ。適用後の読み返し検証なし → `-Verified` 未渡し）
**サブスクリプト**: なし（C# 型 `ResolutionHandler` を `Add-Type` で同居）

## 目的
ディスプレイ解像度を Win32 API（`ChangeDisplaySettings` / `EnumDisplaySettings`）経由で
即時変更するモジュール。レジストリ書き込みと違って再起動なしに反映される（API が
`DISP_CHANGE_RESTART` を返した場合のみ警告）。CSV に複数候補を書いておけるが、通常運用は
Enabled=1 を 1 行にしておく形（解像度の連続切替は副作用が大きいため）。

## 入力 (CSV)
`resolution_list.csv`
- `Enabled`: 有効フラグ（1=実行 / 0=スキップ）
- `Width`: 横幅（ピクセル、整数）
- `Height`: 高さ（ピクセル、整数）
- `Description`: 説明（表示用、例 "Full HD" / "QHD"）
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 1920x1080（Full HD, Enabled=1）, 2560x1440, 1680x1050, 1600x900,
1366x768, 1280x1024, 1280x720（いずれも Enabled=0）。

## 主要ステップ
1. `Add-Type` で C# `ResolutionHandler` 構造体（DEVMODE / 各種定数）をロード
2. `resolution_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
3. `EnumDisplaySettings(ENUM_CURRENT_SETTINGS)` で現在解像度を取得・表示
4. ターゲット一覧をドライラン表示（現在値と一致するエントリは `[SKIP]`）
5. 全件 already-set なら Skipped で early return / 差分があれば実行確認（AutoPilot 自動 Y）
6. 各エントリに `ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY)` を発行
7. 戻り値で分岐: `DISP_CHANGE_SUCCESSFUL` → 成功カウント＋以後の比較用に `currentW/H` を
   更新 / `DISP_CHANGE_RESTART` → 「再起動後に反映」警告して成功扱い / `DISP_CHANGE_FAILED`
   → 失敗 → `New-BatchResult`（`-Verified` は渡さない）

## 注意点・運用メモ
- ハードウェア・ドライバ非対応の解像度は `DISP_CHANGE_FAILED` で失敗扱い
- マルチモニター環境では現状第一プライマリディスプレイのみ対象（device 名 null 指定）
- リフレッシュレートは渡していないため OS 既定が採用される
- 連続変更は前段の成功値を `currentW/H` に反映するため、CSV に複数 Enabled=1 を並べると
  「次に Enabled=1 のものに切り替わる」だけの動作になる（最終的に CSV 末尾に近い解像度が残る）
- 履歴の Verified 列は空欄になる（後述の理由）

## 検証
事前判定としては `GetCurrentResolution()` で現在値と目標値が一致する場合に `[SKIP]` を出す
冪等性チェックがあるが、`ChangeDisplaySettings` 後にもう一度読み返して
期待値一致を確認する Step 5.5 は実装されていない。
これは `DISP_CHANGE_RESTART` のように「呼び出しは成功したが反映は再起動後」の状態を
読み返しでは区別できないこと、および即時反映を保証しても次に新しい接続イベントで
解像度がリセットされうるディスプレイ実装が存在することによる設計判断。
このため `New-BatchResult` 呼び出しに `-Verified` は渡されず、実行履歴の Verified 列は空欄。
