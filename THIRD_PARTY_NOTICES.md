# Third-Party Notices

本リポジトリは、独自実装のキッティングフレームワーク本体（[LICENSE](LICENSE) 参照、MIT）に加え、以下のサードパーティ製ソフトウェアを**バイナリとして同梱**しています。各コンポーネントの著作権表示・ライセンス条件は本ファイルおよび [LICENSES/](LICENSES/) ディレクトリ配下のライセンス全文に従います。

---

## 7-Zip

| 項目 | 内容 |
|---|---|
| 名称 | 7-Zip |
| バージョン | 25.01 (x64), 2025-08-03 リリース |
| 著作権者 | Copyright (C) 1999-2025 Igor Pavlov |
| 公式サイト | https://www.7-zip.org/ |
| ソースコード入手先 | https://www.7-zip.org/download.html |
| 同梱ファイル | `modules/standard/printer_driver_config/tools/7z.exe`<br>`modules/standard/printer_driver_config/tools/7z.dll` |
| 同梱ファイル SHA-256 | `7z.exe`: `4cd7d776c686427226a151789d2d61f0b2ed2c392148cc4e69c0238362fafecf`<br>`7z.dll`: `5bd20fb38499d95c39594f41d4781b6181b3304b7f1f4d06b0182f514e7eaa74` |
| 主ライセンス | GNU LGPL v2.1 or later（[LICENSES/LGPL-2.1.txt](LICENSES/LGPL-2.1.txt)） |
| 構成ライセンス | GNU LGPL v2.1+ / BSD 3-clause / BSD 2-clause / unRAR license restriction の組み合わせ。詳細は [LICENSES/7-Zip-license.txt](LICENSES/7-Zip-license.txt) 参照（7-Zip 公式ライセンスファイル全文） |
| 用途 | `printer_driver_config` モジュールが INF/ 配下の `.exe` / `.zip` 形式のドライバーパッケージを冪等に展開するために使用。fabriq 本体および他モジュールからは使用していない |
| 改変有無 | **未改変**（公式配布のバイナリをそのまま同梱） |

### unRAR 制限について

7-Zip の RAR デコード部分には、原典 unRAR のライセンス条件として「unRAR ソースを RAR 圧縮アルゴリズムの再実装に使用してはならない」という追加制限が課されています。本リポジトリで同梱している `7z.exe` / `7z.dll` を再配布する場合、この制限も併せて伝播します。詳細は [LICENSES/7-Zip-license.txt](LICENSES/7-Zip-license.txt) の "unRAR license restriction" セクションを参照してください。

### ソースコード入手の保証

GNU LGPL v2.1 第 6 条に基づき、本リポジトリは 7-Zip のソースコードを上記公式サイトを通じて入手可能であることを表明します。本リポジトリの管理者は 7-Zip ソースコードの一次配布元ではないため、ソース提供義務は公式サイトへの参照案内をもって履行されます（LGPL v2.1 が認める "written offer" の代替として、永続的な公開ダウンロード URL の提示で要件を満たします）。

---

## 表示・配布の運用

- 本リポジトリのクローンを再配布する際は、`THIRD_PARTY_NOTICES.md` および `LICENSES/` ディレクトリ全体を含めてください
- バイナリのみを切り出して別ロケーションに配置する際は、`modules/standard/printer_driver_config/tools/README-license.txt` も同じディレクトリに保持してください（attribution の伝播）
