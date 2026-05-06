# autologon_config (Standard)

**カテゴリ**: System
**メニュー名**: AutoLogon Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（実装すると平文パスワード露出 / AutoLogonCount 自動デクリメントのため意図的に非実装）
**サブスクリプト**: なし

## 目的
キッティング工程で「再起動を挟んでも一度だけ自動ログオンしたい」需要に応えるため、`HKLM\...\Winlogon` に `AutoAdminLogon` / `DefaultUserName` / `DefaultPassword` / `DefaultDomainName` / `AutoLogonCount=1` を書き込みます。`AutoLogonCount=1` を使うことで「次回 1 回限り」の挙動になり、ログオン後に Windows 自身が資格情報を自動消去するため、放置時のセキュリティリスクを最小化しています。Profile 実行時は `__AUTO_to_<User>__` マーカー経由で `$env:FABRIQ_AUTOLOGON_USER` が渡され、対象ユーザーが自動選択されます。

## 入力 (CSV)
`autologon_list.csv` の主な列:
- `Enabled` … 1=対象 / 0=スキップ
- `No` … 識別番号（対話 No 入力 / Profile 指定でも使用）
- `User` … ユーザー名
- `Password` … パスワード（**`ENC:` プレフィックス暗号化対応**。マスターパスフレーズで復号）
- `Domain` … ドメイン名（ローカルユーザーは空欄）
- `Description` … 表示ラベル
- `Segment` … Segment フィルタ

## 主要ステップ
1. `autologon_list.csv` を読み込み（`Import-ModuleCsv -FilterEnabled`）
2. `$env:FABRIQ_AUTOLOGON_USER` または対話 `No` 番号で対象ユーザー決定（Enabled=1 が 1 件なら自動選択）
3. 対象設定を表示（パスワードはマスク表示）
4. **冪等性チェック**: 現在のレジストリ値（`AutoAdminLogon=1` / `DefaultUserName` 一致 / `AutoLogonCount>=1`）を読み返し、一致なら Skipped で終了
5. 実行確認（AutoPilot 時は自動 Y）
6. レジストリ書き込み（5 値を `Set-ItemProperty -ErrorAction Stop` で一括設定）

## 注意点・運用メモ
- **管理者権限必須**（HKLM 書き込み。明示チェックはせず HKLM 書き込み失敗で検出）
- パスワードは `ENC:` プレフィックス対応。ENC 復号失敗時は対話モードで入力を求める／Error 終了
- `AutoLogonCount=1` を必ず付けるため、無人放置で何度も自動ログオンされる事故を防止
- ローカルアカウント運用時は `Domain` を空欄、ドメイン参加端末でドメインユーザー指定時は `Domain` を埋めること
- Profile 連携時は profile 側に `__AUTO_to_<User>__` を書き、kernel ランナーが `$env:FABRIQ_AUTOLOGON_USER` を立ててからこのモジュールを起動する流れ

## 検証
Post-Apply Verification は **意図的に非実装**:
1. `DefaultPassword` の比較は平文パスワードをログ／メモリに露出しうるため不可
2. `AutoLogonCount` は次回ログオン時に Windows 側で自動デクリメントされる一時値で、適用直後は 1、ログオン後は 0 のどちらも正常状態となり判定が曖昧
3. その代替として、Step 6 で `Set-ItemProperty -ErrorAction Stop` を使い `$failCount` をカウントすることで「書き込みが受理されたか」の実質検証としています

実際の自動ログオン挙動は次回再起動時の動作で確認する運用です（fabriq 標準では `restart_config` 経由）。
