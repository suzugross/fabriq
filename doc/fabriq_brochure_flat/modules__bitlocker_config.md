# bitlocker_config (Standard)

**カテゴリ**: Security
**メニュー名**: BitLocker Enable / BitLocker Disable / BitLocker Await
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（暗号化／復号化は非同期で完了が後続するため、`bitlocker_await` で完了確認する設計）
**サブスクリプト**:
- `bitlocker_config.ps1` … BitLocker 有効化 + 回復キーをエビデンス保存（TPM-only / TpmAndPin 両対応）
- `bitlocker_disable.ps1` … BitLocker 無効化（復号開始）。FullyDecrypted は冪等スキップ
- `bitlocker_await.ps1` … Encrypting / Decrypting 中ドライブの完了をポーリング待機（30 秒間隔、30 分進行なしでタイムアウト）

## 目的
TPM 搭載 PC のキッティングで多用する「BitLocker 有効化 → 完了待機 → 必要なら無効化」を 3 スクリプトに分離して提供するモジュールです。回復キーは `evidence/bitlocker/<日付>_<UID>_<PC名>/<PC名>_<ドライブ名>.txt` に自動保存され、納品時のエビデンスとしてそのまま使えます。TPM-only と TPM+PIN（TpmAndPin Protector）の両対応で、PIN は ENC: プレフィックス暗号化または `hostlist.csv` 経由 (`$env:SELECTED_PIN`) で安全に注入可能です。PIN が記号／英字を含む場合は `UseEnhancedPin=1` を自動設定します。

## 入力 (CSV)
`bitlocker_list.csv` の主な列:
- `Enabled` … 1=対象 / 0=スキップ
- `TargetDrive` … 対象ドライブ（例: `C:`, `D:`）
- `EncryptionMethod` … `XtsAes128` / `XtsAes256` 等
- `UsedSpaceOnly` … TRUE/FALSE（使用領域のみ暗号化）
- `SkipHardwareTest` … TRUE/FALSE
- `AutoUnlock` … TRUE/FALSE（データドライブ向け）
- `Pin` … TpmAndPin 用 PIN（空欄=TPM-only）。**`ENC:` プレフィックス対応**
- `Description` / `Segment`

PIN 解決優先順: `$env:SELECTED_PIN` > `Pin` 列 > なし（TPM-only）

## 主要ステップ
[Enable]
1. TPM 状態確認（`TpmPresent` 必須、`TpmReady=false` は Warning）
2. `bitlocker_list.csv` 読み込み（Enabled=1）
3. PIN 解決＋ENC: ガード判定
4. ドライブ存在確認＋現状表示
5. 実行確認（AutoPilot 時は自動 Y）
6. 必要なら FVE レジストリポリシー（`UseAdvancedStartup=1` / `UseTPMPIN=1` / `UseEnhancedPin=1`）設定
7. エビデンス出力先決定 → ドライブごとに `Enable-BitLocker` 実行 → 回復キー取得 → エビデンス書き出し → AutoUnlock 適用

[Disable] CSV 読込 → 状態判定 → 実行確認 → `Disable-BitLocker`
[Await]   CSV 読込 → Encrypting/Decrypting 検出 → 実行確認 → 30 秒ポーリング → 30 分進行なしで個別タイムアウト

## 注意点・運用メモ
- **管理者権限必須**（3 スクリプト共通）
- 回復キーは **必ず Evidence に保存**（TpmPin 内容は記録しない、セキュリティ配慮）
- AutoUnlock はシステムドライブが FullyEncrypted になってから有効化される仕様のため、Enable 直後の AutoUnlock 適用は失敗しがち（Warning のみで継続）。Await 通過後の再実行で適用される
- FVE レジストリポリシーを書き換えるため、グループポリシー管理環境では運用方針との衝突を事前確認
- Profile 構成例: `... → bitlocker_config.ps1 → bitlocker_await.ps1 → ...`

## 検証
3 スクリプトいずれも `-Verified` フラグは未返却。理由は「BitLocker の状態遷移は本質的に非同期」であり、Enable/Disable 直後に Verified=true を返してしまうと「完了」を意味してしまい誤解を招くためです。完了判定は `bitlocker_await` の責務に分離されています。Enable は「回復キー取得 + エビデンス保存」をもって受理成功扱い、Disable は `Disable-BitLocker` 受理をもって成功扱い、Await は `FullyEncrypted` / `FullyDecrypted` で集計します。

冪等性: Enable は `ProtectionStatus=On` のドライブを Skip（PIN 追加が必要なら Tpm→TpmAndPin にアップグレード）、Disable は `FullyDecrypted` / `DecryptionInProgress` を Skip。
