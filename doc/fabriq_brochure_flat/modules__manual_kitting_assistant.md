# manual_kitting_assistant (Extended)

**カテゴリ**: ManualWorks
**メニュー名**: Manual Kitting Assistant
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（手動操作の結果は OS 側の個別モジュールで検証する設計分離）
**サブスクリプト**: なし（`prompt/` フォルダに手順テキストファイル `.txt` を別途配置）

## 目的
レジストリ/CLI で自動化できない手動キッティング作業を、画面右下に常駐する WinForms モーダル風アシスタントで案内するモジュール。
1 ステップ = 1 画面（Step ID, タイトル, 手順説明, 実行ボタン, コピー1/2/3, 貼り付け, 完了）の構成で、operator はボタンクリックでアプリ起動・クリップボードコピー・貼り付けを行いながら次へ進む。
GUI ポリシーは「**マウスのみ**、F2 キーのみ完了に割当」「クリックしてもフォーカスを盗まない (`WS_EX_NOACTIVATE` + `WM_MOUSEACTIVATE` 2 段ガード)」「Ctrl+V は C# `keybd_event` ラッパー経由で送信し PowerShell マーシャリング問題を回避」。
カラーテーマは「Gundam Light」（mecha white 背景 + federation blue ヘッダー + V-fin gold アクセント + 各種ガンダムカラーボタン）。

## 入力 (CSV)
`step_list.csv`:
- **Enabled**: 有効フラグ
- **StepID**: ステップ ID（例 `S001`）
- **StepTitle**: GUI ヘッダー表示名
- **PromptFile**: `prompt/` 配下の手順説明テキストファイル名（空欄ならコンパクトモード = 説明エリアなし）
- **OpenCommand**: 実行ボタンで起動するコマンド
- **OpenArgs**: コマンド引数
- **Copy1** / **Copy2** / **Copy3**: 各コピーボタンでクリップボードに送る値
- **Segment**: Segment フィルタ

## 主要ステップ
1. `System.Windows.Forms` / `System.Drawing` ロード + C# `FabriqCtrlCSender` / `NoActivateButton` / `NoActivateForm` を `Add-Type`
2. CSV 読み込み + `prompt/` ディレクトリ存在確認 + 全 PromptFile の実在検証（不足あれば事前 Error 終了）
3. `Confirm-ModuleExecution`
4. WinForms ウィンドウを画面右下に配置（modeless, TopMost）
5. メインループ: 各ステップで `Update-StepDisplay` で UI 更新 → `DoEvents` ポーリング待機 → 完了 or キャンセル待ち
6. キャンセル時は MessageBox で確認、Yes なら以降 Skip 集計
7. 全ステップ完了後 `New-BatchResult` 集計

## 注意点・運用メモ
- 対話型セッション必須（GUI 描画必要）
- AutoPilot 運用には不向き（GUI 表示後は operator のクリック待ちでブロックされる）。AutoPilot プロファイルからは除外推奨
- PromptFile が空のステップはコンパクトモード（フォーム高さを動的縮小）で表示
- 「貼り付け」ボタンは `GetForegroundWindow` で押下時のフォアグラウンドウィンドウを記憶 → `SetForegroundWindow` でフォーカス戻し → `keybd_event` で Ctrl+V を送出する 3 段構成
- コピー成功時はボタンを 600ms グリーンフラッシュさせる視覚フィードバック
- 全ボタンに `NoActivateButton` を使い、フォーム自体も `WS_EX_NOACTIVATE` でアクティブ化を抑制（操作中の他ウィンドウのフォーカス維持）

## 検証
未実装。手動操作の結果検証は OS 側の個別モジュール（`reg_hkcu_config` 等）の責務とする設計分離。`-Verified` 未渡しで Verified 列は空欄。
