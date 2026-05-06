# robocopy_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: Robocopy
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（ExitCode 判定のみ。実ファイル一致検証なし、`-Verified` 未渡し）
**サブスクリプト**: なし

## 目的
CSV (`robocopy_list.csv`) に列挙したジョブを `robocopy.exe` で順次実行するモジュール。
主要オプション (Recursive / Mirror / CopyACL / SkipOlder) を 0/1 フラグ列で表現してタイポ事故を防ぎ、
`/MT` `/Z` `/XD` などの追加引数は CustomOptions 列で自由記述する設計。
UNC 認証（`net use`）を内蔵しており、SMB 共有を含むコピーも 1 ジョブで完結させられる。
ENC: プレフィックスによる暗号化パスワードに対応。

## 入力 (CSV)
`robocopy_list.csv`
- `Enabled`: 1=有効 / 0=無効
- `ID`: ジョブ識別子
- `Source`, `Destination`: コピー元・先パス（ローカル / UNC）
- `Recursive`: 1=`/E`
- `Mirror`: 1=`/MIR`（Recursive と両立時は Mirror 優先）
- `CopyACL`: 1=`/COPYALL /DCOPY:DAT` / 0=`/COPY:DAT /DCOPY:DAT`
- `SkipOlder`: 1=`/XO`
- `FileFilter`: ファイル名 / ワイルドカード（空欄=全ファイル、スペース区切り複数指定可）
- `CustomOptions`: 追加オプション自由記述
- `AuthUser`, `AuthPass`: UNC 認証情報（任意、ENC: 対応）
- `Description`, `Segment`

## 主要ステップ
1. `robocopy_list.csv` を `Import-ModuleCsv -FilterEnabled` で読み込み
2. 各ジョブのドライラン表示（オプション展開、AuthUser はマスク）
3. 実行確認（AutoPilot 自動 Y）
4. ジョブループ:
   - 4-1. AuthUser/AuthPass 両方ありなら `net use \\Server\Share /user:...` で接続
   - 4-2. `robocopy Source Destination [FileFilter] /R:3 /W:5 /NP <flags>` 実行（baseline オプション強制付与）
   - 4-3. ExitCode 評価 (0=変更なし / 1=コピー成功 / 2-3=余剰検出 / 4-7=ミスマッチ Warning / 8+=Error)
   - 4-4. finally で `net use /delete` 切断
5. 集計して `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（CopyACL=1 で SeSecurityPrivilege/SeBackupPrivilege が必要）
- baseline オプション `/R:3 /W:5 /NP` は CSV から除外不可（ハングアップ防止のセーフガード）
- パスワードはコンソール / トランスクリプトに一切出力しない（ログ漏洩対策）
- Mirror=1 は Source 不在ファイルを Destination から削除するため、誤 Source で意図せぬ削除事故の
  リスクあり。初回テストは Recursive=1 + SkipOlder=1 を推奨
- Mirror=1 は厳密な冪等ではない（Source 変化に追従するため）
- `FileFilter` は Source をディレクトリ指定のままにしてファイル名はこの列で渡す形式
- net use 失敗時はジョブ Skip して次へ進行（ジョブ間隔離）
- robocopy ExitCode 仕様の Warning レンジ (4-7) は Success 扱い

## 検証
robocopy の ExitCode は「ジョブが成功したか」を返すだけで、実ファイル単位の一致は保証しない。
本モジュールは実ファイル比較や ACL 比較の Step 5.5 を持たず、`New-BatchResult` に `-Verified` を
渡していないため履歴の Verified 列は空欄。手動検証としては `robocopy Source Destination /L /E /NP`
（`/L` でドライラン）や `icacls` での ACL 確認が Guide で案内されている。
将来的に DFS-R や Hash 比較を組み込むには copyfile_config 系の検証除外パターン
（`project_verification_exclusions.md` 参照）と整合を取る必要がある。
