# acl_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: ACL Backup / ACL Restore
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 不可（誤 PASS リスクのため意図的に非実装）
**サブスクリプト**:
- `acl_backup.ps1` … `icacls /save /T` によるディレクトリ ACL のフルバックアップ + 継承断ち subdir の個別バックアップ
- `acl_restore.ps1` … フル復元 → 浅い順に個別 override の二段復元

## 目的
キッティング作業や移行作業に伴ってファイル ACL を巻き戻す必要が出た場合に備え、対象ディレクトリツリーの ACL をハイブリッド方式（フル `/save /T` ＋ 継承断ち subdir の個別保存）でバックアップ／レストアします。フルセーブのみでは継承断ちフォルダの ACL がわずかにずれ、サブディレクトリを 1 つずつ保存すると大規模ツリーで激遅になるという `icacls` の二律背反を、ハイブリッド方式で解消します。冪等性の観点では、復元時はフル → 浅い順 → 深い順という順序を守ることでより具体的な ACL が必ず後勝ちで適用される設計になっています。

## 入力 (CSV)
`acl_list.csv` の主な列:
- `Enabled` … 1=処理 / 0=スキップ
- `Id` … バックアップフォルダ名（`backup/{Id}_{SafeName}/`）に使う一意識別子
- `TargetPath` … バックアップ／復元の起点ディレクトリ。`%USERPROFILE%` `%SELECTED_NEW_PCNAME%` 等の環境変数を `Expand-UserEnvironmentVariables` で展開
- `Description` … 表示ラベル
- `Segment` … `Import-ModuleCsv` による Segment フィルタ用（空欄は常に採用）

## 主要ステップ
1. `acl_list.csv` 読み込み（`Import-ModuleCsv -FilterEnabled` で Segment フィルタ込み）
2. Pre-flight: `Test-AdminPrivilege` と `icacls.exe` 存在確認
3. Pre-execution display（dry-run 一覧）
4. 実行確認（AutoPilot 時は自動 Y）
5. Backup: フル `/save /T` を撮ったあと、継承断ち subdir を SHA256 ハッシュ名で個別保存し `_manifest.csv` を出力 / Restore: フル復元 → manifest を浅い順にソートして個別 override
6. 結果集計（`New-ModuleResult`）

## 注意点・運用メモ
- **管理者権限必須**（`icacls /save` `/restore` ともに admin token が必要）
- バックアップ出力構造は `backup/{Id}_{SafeName}/_full_acl.txt`, `_manifest.csv`, `individual/{hash}_acl.txt`
- シンボリックリンク／ジャンクションはスキャン対象外
- Restore は事前 Backup が無いと Error
- 継承断ち subdir が無い場合はフル復元のみで完了
- `_manifest.csv` がパス→ハッシュ名のマッピング表として残るので、ハッシュファイルから元パスへ逆引き可能

## 検証
本モジュールは Post-Apply Verification を **意図的に非実装** としています。これは fabriq プロジェクトルール（CLAUDE.md「Post-Apply Verification 除外モジュール」に準拠）の判断で、`icacls /save` のテキスト出力と現在の ACL を厳密に比較するには継承フラグ／SID 解決／ACE 順序／差分検出を完全に再現する必要があり、誤 PASS（false positive）リスクが過大なためです。「PASS と表示されたが実際は不一致」という最悪ケースを誘発しないよう、`-Verified` を渡さず Verified 列は空欄になります。検証は `icacls /save` の出力テキストを目視差分するか外部ツールで実施することが推奨されます。
