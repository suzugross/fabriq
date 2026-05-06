# ssid_config (Standard)

**カテゴリ**: Network
**メニュー名**: SSID Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（プロファイル一覧再列挙で存在確認）
**サブスクリプト**: なし（XML を実行時に動的生成し %TEMP% 経由で `netsh` に渡す）

## 目的
Wi-Fi (WLAN) プロファイルを `netsh wlan add profile` で登録するモジュール。
CSV に定義した SSID ごとに WLAN プロファイル XML を動的生成し、`%TEMP%` に UUID 付きで
書き出して netsh に取り込ませる。WPA2PSK / WPA3SAE / open の 3 認証方式に対応。
パスフレーズは平文だけでなく `ENC:` プレフィックス付きの暗号化文字列にも対応し、
Fabriq Studio で暗号化したものを実行時に自動復号する。

## 入力 (CSV)
`ssid_list.csv`
- `Enabled`: 1=登録 / 0=スキップ
- `SSID`: ネットワーク名
- `Authentication`: `WPA2PSK` / `WPA3SAE` / `open`
- `Encryption`: `AES` / `none`（open は none）
- `Password`: パスフレーズ（open は空欄、`ENC:` 暗号化対応）
- `AutoConnect`: 1=auto / 0=manual
- `NonBroadcast`: 1=隠れネットワーク / 0=通常
- `Description`, `Segment`

## 主要ステップ
1. `ssid_list.csv` 読み込み（Enabled=1 のみ）
2. 前提チェック: `netsh wlan show profiles` で WLAN サービス応答 + 既存プロファイル名抽出
   （日本語版・英語版両方の出力に対応するパース）
3. ドライラン表示（既存プロファイルと大文字小文字無視で照合 → `[SKIP]` / `[ADD]`）
4. 実行確認（AutoPilot 自動 Y）
5. 適用ループ:
   - 既登録は Skip（冪等性、上書きはしない）
   - Authentication に応じた XML を動的生成（open は sharedKey ブロック無し、
     WPA2PSK/WPA3SAE は passPhrase 平文埋め込み）
   - SSID / Password は `[SecurityElement]::Escape` で XML エスケープ
   - 一時 XML を `%TEMP%\<UUID>.xml` に生成 → `netsh wlan add profile filename=...` 実行
   - finally 句で一時 XML を必ず削除（パスワード残留防止）
6. Step 5.5: Post-Apply Verification（再度 `netsh wlan show profiles` 実行 → 全対象 SSID が
   プロファイル一覧に存在するか検証）→ `New-BatchResult -Verified $verified` 返却

## 注意点・運用メモ
- 管理者権限必須
- WLAN サービス無効・無線アダプタ無しの環境では前提チェックで失敗
- 既存プロファイルは上書きしない設計（パスワード変更等は事前に `netsh wlan delete profile` 必要）
- WPA3SAE は Windows 10 2004 以降 + 対応アダプタのみ機能。非対応環境では登録は成功するが
  接続時にエラーになる可能性
- `ENC:` 復号失敗時は当該行のみエラーで他は継続
- パスワードがコンソール / トランスクリプトに出力されないようマスキング徹底
- 一時 XML は中断時でも finally で削除（パスワード残留リスク最小化）

## 検証
Step 5.5 では `netsh wlan show profiles` を再実行して出力をパースし、
登録対象 SSID が全件プロファイル一覧に含まれているかを大文字小文字無視で照合。
1 件でも欠損があれば `-Verified=$false` で `New-BatchResult` に渡す。
ただし「プロファイルが存在する」ことの検証であり、「実際にその AP に接続できる」ことまでは
保証しない（電波状況や AP 側設定は対象外）。プロファイルの中身（パスワード値等）も読み返し
比較していない（`netsh wlan show profile name=X key=clear` を呼ぶことになりログに平文が
残るため意図的に避けている）。
