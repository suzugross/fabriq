# group_config (Extended)

**カテゴリ**: User Management
**メニュー名**: Local Group Member Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（実装可能なリファレンス記述あり、`Test-LocalGroupMemberExists` を Step 5.5 で再呼び出しすれば Verification 化できる構造）
**サブスクリプト**: なし

## 目的
ローカルグループ（Administrators, Remote Desktop Users 等）にドメイングループ／ドメインユーザー／ローカルユーザー／現在のログインユーザーをメンバーとして追加するモジュール。
特徴は `MemberType=CurrentUser` のサポートで、UAC 昇格時でも `Get-CimInstance Win32_ComputerSystem` の `UserName` から実際の対話セッションユーザーを正しく解決する（昇格前のユーザーが返る）。

## 入力 (CSV)
`group_list.csv`:
- **Enabled**: 有効フラグ
- **LocalGroup**: 追加先ローカルグループ名
- **MemberType**: `DomainGroup` / `DomainUser` / `LocalUser` / `CurrentUser`
- **MemberName**: 追加メンバー名（CurrentUser では空欄）
- **Domain**: ドメイン名（LocalUser/CurrentUser では空欄）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み
3. ドライラン表示（各行で `Test-LocalGroupExists` + `Test-LocalGroupMemberExists` を呼び `[Current]` / `[Change]` / `[ERROR]` マーカー）
4. `Confirm-ModuleExecution`
5. メンバー追加ループ（`Build-MemberName` で `Domain\User` 形式整形 → 冪等性チェック → `Add-LocalGroupMember`）
6. `New-BatchResult` 集計

## 注意点・運用メモ
- ドメインメンバー追加にはドメイン接続が必要
- 冪等性ロジックは MemberType により分岐: LocalUser は `COMPUTERNAME\Name` または `Name` の完全一致、それ以外は `*\Name` の末尾一致で判定
- CurrentUser 解決は WMI 失敗時、`USERDOMAIN ≠ COMPUTERNAME` で `USERDOMAIN\USERNAME`、それ以外で `USERNAME` にフォールバック

## 検証
未実装だが、Guide.txt 内に「Step 5.5 で `Test-LocalGroupMemberExists` を再呼び出しすれば実装可能」と明記されており構造的には準備されている。`-Verified` 未渡しで Verified 列は空欄。
