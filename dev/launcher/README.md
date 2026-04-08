# Fabriq Launcher

`Fabriq.exe` を生成するための C# 極小ランチャーのソース一式です。

## 目的

- `Fabriq.bat` と同じく `kernel\main.ps1` を起動するだけのラッパー
- **カスタムアイコン** と **UAC 自動昇格マニフェスト** を埋め込み、エクスプローラー・タスクマネージャー上で「アプリ」として見せる
- fabriq 本体（`kernel/`, `modules/`, `apps/` 等）には一切手を加えない
- 依存は **Windows 標準の .NET Framework 4.x 付属 `csc.exe` のみ**

## ファイル構成

| ファイル | 役割 |
|---|---|
| `Launcher.cs` | C# ソース。自身のディレクトリを cwd にして `conhost + powershell -File kernel\main.ps1` を起動 |
| `app.manifest` | UAC `requireAdministrator` を指定。ダブルクリック時に直接 UAC ダイアログが出る |
| `fabriq.ico` | アイコン（初回ビルド時に `shell32.dll` から仮アイコンを自動抽出） |
| `build.ps1` | ビルドスクリプト。`csc.exe` を呼び出して `..\..\Fabriq.exe` を生成 |

## ビルド手順

PowerShell で以下を実行するだけです（**管理者権限は不要**）。

```powershell
cd e:\fabriq\dev\launcher
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

成功すると `e:\fabriq\Fabriq.exe` が生成されます。

## アイコンの差し替え

1. 任意の `.ico` ファイルを `dev\launcher\fabriq.ico` として上書き保存
2. `build.ps1` を再実行

`.ico` 生成が手元にない場合は、PNG から以下のようなオンラインツールで変換するか、`ImageMagick` の `magick convert input.png -define icon:auto-resize=256,128,64,48,32,16 fabriq.ico` などで作成してください。

## 再ビルドが必要になるタイミング

以下のいずれかを変更した場合に再ビルドが必要です：

- `Launcher.cs` （製品名・バージョン・起動ロジック）
- `app.manifest` （UAC 設定・対応OS）
- `fabriq.ico` （アイコン差し替え）

fabriq 本体（`kernel/`, `modules/`, `main.ps1` など）を変更しても**再ビルドは不要**です。ランチャーはあくまで `kernel\main.ps1` を実行するだけのラッパーなので、main.ps1 側の変更はそのまま反映されます。

## 既存の Fabriq.bat との関係

`Fabriq.bat` はそのまま**残されています**。以下のような場合に使い分けてください：

| 用途 | 推奨エントリーポイント |
|---|---|
| 通常起動（配布先・本番） | `Fabriq.exe` |
| デバッグ・コンソール出力を追いたい | `Fabriq.bat` |
| ランチャーがビルドされていない環境 | `Fabriq.bat` |

## トラブルシューティング

### `csc.exe` が見つからないと言われる
`.NET Framework 4.x` がインストールされていません。Windows 10/11 なら標準で入っているはずですが、極端に古い環境では `Windows Features` から「.NET Framework 4.x Advanced Services」を有効にしてください。

### ビルドで "file is in use" エラー
`Fabriq.exe` が実行中です。タスクマネージャーでプロセスを終了してから再ビルドしてください。

### SmartScreen の警告が出る
`Fabriq.exe` は未署名のため、初回実行時に SmartScreen 警告が出る場合があります。「詳細情報」→「実行」で起動できます。恒常的に警告を回避するにはコード署名証明書が必要です。
