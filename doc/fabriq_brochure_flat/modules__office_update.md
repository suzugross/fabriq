# office_update (Standard)

**カテゴリ**: Maintenance
**メニュー名**: Office Update
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: スクリプト内で VersionToReport 読み返し（事実上の検証、`-Verified` は未渡し）
**サブスクリプト**: `office_update.ps1`（メイン処理 1 本のみ）

## 目的
Click-to-Run（C2R）版 Office の更新を実行し、完了を待機する専用モジュール。
`OfficeC2RClient.exe /update` を起動し、バージョンレジストリ変化・
`Scenario` レジストリ・OfficeC2RClient プロセス存在の 3 シグナルを組み合わせた
ハイブリッド検知で完了を判定します。MSI 版（2016 以前）は対象外で、
Click-to-Run 構成が検出できない場合は Skip します。

## 入力 (CSV)
`office_update_list.csv`（SettingName/Value ペア形式・Segment 列なし）:
- `Enabled`: 行有効フラグ
- `SettingName`: 設定項目名（下記 4 種）
- `Value`: 設定値
- `Description`: 説明

採用される SettingName:
- `TimeoutMinutes`: 更新完了までの最大待機分（既定 60）
- `PollIntervalSeconds`: ポーリング間隔秒（既定 10）
- `ForceAppShutdown`: 1=Office アプリを事前強制終了
- `DisplayLevel`: 1=更新 UI 表示

## 主要ステップ
1. `office_update_list.csv` から設定読込み（SettingName/Value をハッシュ化）
2. C2R 設定レジストリ確認、未インストールなら Skip
3. `OfficeC2RClient.exe` 存在確認 + 現バージョン取得 + `Wait-NetworkReady`
4. ドライラン表示（製品 / バージョン / チャネル / 各設定値）
5. 実行確認（AutoPilot は自動 Y）
6. ForceAppShutdown=1 なら Office プロセス（WINWORD, EXCEL, POWERPNT, OUTLOOK,
   ONENOTE, MSACCESS, MSPUB, VISIO）を `Stop-Process -Force`
7. `OfficeC2RClient.exe /update user displaylevel=... forceappshutdown=...` を `Start-Process`
8. 完了待機ループ:
   - シグナル 1: `VersionToReport` レジストリ変化 → 確定的に完了
   - シグナル 2: `OfficeC2RClient` プロセス存在 → 更新処理中
   - シグナル 3: `HKLM\...\ClickToRun\Scenario` サブキー存在 → 更新操作中
   - 30 秒以上経過 + プロセス無 + Scenario 無 → 処理完了
9. `VersionToReport` 読み返し: 変化あり=Success / タイムアウト=Error / 早期完了で変化なし=Skipped

## 注意点・運用メモ
- C2R 版 Office の更新は通常再起動不要
- `OfficeC2RClient.exe` のコマンドラインスイッチは Microsoft 非公式
- ForceAppShutdown=1 では未保存ドキュメントが失われるため、
  Profile での自動実行時は事前注意喚起が必要
- 完了検知の min wait（30 秒）は更新起動直後の false negative を防ぐためのバッファ

## 検証
スクリプト内で `VersionToReport` レジストリ値の事前/事後比較を行い、
バージョン変化の有無で Status / Message を決定。`-Verified` 引数は未渡しのため
実行履歴の Verified 列は空欄になるが、事実上の検証はスクリプト内で完結。
将来 `-Verified` を追加するなら `(afterVersion -ne beforeVersion)` を直接渡せる構造。
