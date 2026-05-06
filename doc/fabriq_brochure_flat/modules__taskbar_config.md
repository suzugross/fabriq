# taskbar_config (Standard)

**カテゴリ**: Desktop
**メニュー名**: Taskbar Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（適用先が Default プロファイル＝将来の新規ユーザーのため `-Verified` 未渡し）
**サブスクリプト**: なし（CSV から `LayoutModification.xml` を動的生成）

## 目的
タスクバーにピン留めするアプリを CSV ベースで定義し、`LayoutModification.xml` を
動的生成して Default User プロファイルへ配置するモジュール。配置後に作成される新規
ユーザーがこの XML を継承する形で、最初からタスクバーに業務用アプリのみが並んだ状態で
セッションを開始できる。生成 XML は `sysprep_config/source/` にも自動コピーされ、
Sysprep ワークフローの素材として 1 度の実行で 2 箇所同時に最新化される設計。

## 入力 (CSV)
`taskbar_list.csv`
- `Enabled`: 1=ピン留め対象 / 0=スキップ
- `Order`: タスクバー上の表示順（昇順ソート）
- `LinkPath`: アプリパス（`.lnk` または `.exe`、`%APPDATA%` 等の環境変数使用可）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: File Explorer (Order=10) / Google Chrome (Order=20)。

## 主要ステップ
1. `taskbar_list.csv` 読み込み → Order 昇順ソート
2. 各 LinkPath の存在確認（実行時点で不在でも Warning のみで処理継続）
3. ドライラン表示
4. 実行確認（AutoPilot 自動 Y）
5. `LayoutModification.xml`（CustomTaskbarLayoutCollection / TaskbarPinList）を動的生成 →
   `C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml` 配置
6. 同 XML を `modules\standard\sysprep_config\source\LayoutModification.xml` に自動コピー
   （無条件・毎回。コピー失敗は警告のみで Success 維持）
7. ファイル存在 + サイズ>0 チェック後 `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（Default User プロファイルへの書き込み）
- 既存 `LayoutModification.xml` は上書き
- LinkPath がキッティング時点で不在でも XML には記載される（後続モジュールでアプリを
  入れる Profile でも問題なし）
- 全エントリ Enabled=0 の場合、空の TaskbarPinList を持つ XML が生成され、
  デフォルトのピン留め（Edge 等）も除去される（確信犯的な無効化用途あり）
- `sysprep_config` の `setupcomplete_list.csv` には対応する `CopyFile` 行が同梱済みで、
  Sysprep 実行時に source/ → Setup\Scripts\source\ → Default User Shell\ への二段ホップで
  最終配置される
- ショートカット格納場所:
  - 全ユーザー共通: `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\`
  - 現在ユーザー: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\`

## 検証
Post-Apply Verification は実装されていない。XML 書き出し直後にファイル存在 + サイズ>0 の
最低限チェックは行うが、XML 内容のパース検証ステップはなし。
そもそも実際にタスクバーに反映されるのは「配置後に新規作成される Default User プロファイル」
であり、現セッションでの読み返し検証は仕組み上困難（現ユーザーには既存タスクバーが
存在するため、本モジュールの効果は次のユーザーアカウント作成時まで観測できない）。
`-Verified` を `New-ModuleResult` に渡していないため履歴の Verified 列は空欄。

実効性確認は新規ユーザー作成 → ログオンしてタスクバーを目視、または Sysprep 後のイメージ
展開先での OOBE 後タスクバー確認。
