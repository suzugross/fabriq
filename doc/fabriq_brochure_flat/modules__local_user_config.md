# local_user_config (Standard)

**カテゴリ**: User Management
**メニュー名**: Create Local Users / Delete Local Users
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: Create のみ実装あり（ユーザー存在 + グループ所属検証）／Delete はなし
**サブスクリプト**: `local_user_config.ps1`（作成）, `local_user_delete.ps1`（削除）

## 目的
ローカルユーザーアカウントの作成・削除を行います。
全 PC 共通の `local_user_list.csv` と PC 固有の `local_user_host_list.csv` を
union 適用するハイブリッド構造で、共通管理者アカウントは前者で一元管理しつつ、
特定機種だけにローカル運用ユーザーを足す等の柔軟な運用が可能です。
グループ所属はセミコロン区切り複数指定（例: `Users;Remote Desktop Users`）。

## 入力 (CSV)
**`local_user_list.csv`（全 PC 共通）**:
- `Enabled`, `UserName`, `Password`
- `PasswordNeverExpires`: 1=無期限, 0=有効期限あり
- `UserMayNotChangePassword`: 1=変更禁止, 0=変更可
- `Group`: 所属グループ（セミコロン区切り複数可）
- `Description`, `Segment`

**`local_user_host_list.csv`（任意・PC 固有）**:
- 上記列に加え `NewPCName`（hostlist の NewPCName と完全一致、SELECTED_NEW_PCNAME で突合）

## 主要ステップ（Create）
1. `local_user_list.csv` 読み込み
2. `local_user_host_list.csv` 存在時、SELECTED_NEW_PCNAME 一致行を append
3. ユーザー一覧表示（無効行は [DISABLED] と表示）
4. 実行確認（AutoPilot は自動 Y）
5. 作成ループ: `Get-LocalUser` で既存検出時 Skip、`New-LocalUser` で作成
   （`AccountNeverExpires=$true` 固定、PasswordNeverExpires / UserMayNotChangePassword は CSV 値反映）
6. グループ追加: `Add-LocalGroupMember` を Group 列各値ごとに実行
7. Step 5.5: Verification → `New-BatchResult -Verified`

## 主要ステップ（Delete）
1. CSV 読み込み（Enabled フィルタなし＝全行が削除候補）
2. host_list union
3. 一覧表示 + 確認
4. `Remove-LocalUser`、不在は Skip
5. `New-BatchResult` で集計（Verified なし）

## 注意点・運用メモ
- 管理者権限必須（`New-LocalUser` / `Remove-LocalUser`）
- パスワードは平文 CSV 保存。crypto レイヤ（`ENC:` プレフィクス + passphrase）対応は
  framework 共通機能で別途提供
- ユーザー作成失敗時もグループ追加は実行されないが、グループ追加失敗だけでは
  `successCount` には影響せず、警告のみ
- Delete スクリプトは Enabled フィルタを使わない（全行対象）。
  CSV を「削除対象リスト」として再利用する思想

## 検証
Create 側のみ Verification 実装。`Get-LocalUser` でユーザー存在を確認し、
`Get-LocalGroupMember` で `Group` 列指定の各グループに対し正しく所属しているかを
正規表現マッチで検証。1 件でも欠落があれば `-Verified $false`。
Delete 側は Verification 未実装で `-Verified` 引数を渡さない。
