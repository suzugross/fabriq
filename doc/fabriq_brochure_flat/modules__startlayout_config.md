# startlayout_config (Standard)

**カテゴリ**: Desktop
**メニュー名**: Start Layout Backup / Build / Import / Delete
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（Backup/Delete はファイル・cmdlet レベル検証あり、`-Verified` 未渡し）
**サブスクリプト**:
- `startlayout_backup_config.ps1` … 現スタートレイアウトを `Export-StartLayout` で JSON 出力
- `startlayout_build_config.ps1` … JSON → customizations.xml → .ppkg を `ICD.exe` でビルド
- `startlayout_import_config.ps1` … `Install-ProvisioningPackage` で PPKG 適用
- `startlayout_delete_config.ps1` … 既インストール PPKG をアンインストール + ファイル物理削除

## 目的
Windows 11 のスタートメニュー（ピン留めレイアウト）を「Backup → Build → Import → Delete」
の 4 フェーズで管理するモジュール群。Backup でひな型を抽出し、Build で Provisioning
Package (.ppkg) に変換し、Import で配布、Delete で解除という ppkg ベースの正規ルートを
1 モジュールに収めている。Win11 の `Export-StartLayout` JSON 形式が前提なので
Win10 環境では非対応。

## 入力 (CSV)
`startlayout_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `Id`: 連番識別子
- `FileName`: 拡張子なしの共通ベース名（`startlayout` 指定時に
  `json/startlayout.json` / `xml/startlayout.xml` / `ppkg/startlayout.ppkg` が入出力対象）
- `Segment`: Segment フィルタ（任意）

サブディレクトリ `json/` `xml/` `ppkg/` は実行時自動作成。

## 主要ステップ
**[Backup]** 1. CSV読込 2. `Export-StartLayout` 存在確認 3. ドライラン (`[NEW]`/`[OVERWRITE]`)
4. 確認 5. JSON 出力 6. ファイル存在 + サイズ>0 検証

**[Build]** 1. CSV読込 2. ICD.exe / StoreFile / 入力 JSON 検出 3. ドライラン (`[NEW]`/`[REBUILD]`)
4. 確認 5. JSON 圧縮 + XML エスケープ → `customizations.xml`（UTF-8 BOM）生成 →
`ICD.exe` で .ppkg ビルド

**[Import]** 1. CSV読込 2. `Install-ProvisioningPackage` cmdlet と PPKG ファイル確認
3. ドライラン (`[INSTALL]`/`[REINSTALL]`) 4. 確認 5. `Install-ProvisioningPackage -QuietInstall -ForceInstall`

**[Delete]** 1. CSV読込 2. `Get-/Remove-ProvisioningPackage` 確認 3. ドライラン
(`[INSTALLED]`+PackageId / `[NOT FOUND]`) 4. 確認 5. Phase 1 cmdlet 削除 → Phase 2 .ppkg 物理削除
→ Phase 3 再クエリで完全解除を検証

## 注意点・運用メモ
- Win11 専用（Export-StartLayout JSON 形式が前提）
- Build は Windows ADK の Imaging and Configuration Designer 必須。
  ICD.exe 検索パスは `${ProgramFiles(x86)}\Windows Kits\10\...` 系を 2 候補 + PATH の順に走査
- Import/Delete は管理者権限必須、Backup は不要
- Delete は PPKG 経由で配布した他の設定（ポリシー等）も同時に解除される副作用あり
- Build に渡す JSON は `pinnedList` キー前提（無くても継続するが Warning）
- 典型運用は「Build は手動、Import を Profile に組み込み」（ICD.exe 環境を全クライアントに
  揃えなくて済むように）

## 検証
- **Backup**: 出力後にファイル存在 + サイズ>0 を内部検証
- **Build**: 中間 XML / 最終 PPKG の生成可否は実装で検証するが `-Verified` 未渡し
- **Import**: `Install-ProvisioningPackage` 戻り値のみ。実際にスタートレイアウトに反映されるかは
  再ログオンが必要なケースあり（その場で読み返し検証は不可）
- **Delete**: Phase 3 で `Get-ProvisioningPackage` 再クエリにより完全削除を確認するが
  `-Verified` 引数は未使用

いずれも `New-BatchResult -Verified` 未渡しのため履歴 Verified 列は空欄。
最終的な「ピン留めが意図通り適用された」確認はサインイン後の目視か Get-StartLayout の
出力比較に委ねる設計。
