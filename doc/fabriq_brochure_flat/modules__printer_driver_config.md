# printer_driver_config (Standard)

**カテゴリ**: Printer
**メニュー名**: Printer Drivers / Register Printers / Uninstall Printer Drivers
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: install=実装あり／register=実装あり（4 項目突合）／uninstall=なし
**サブスクリプト**: `printer_driver_install.ps1`（ドライバ登録）, `printer_config.ps1`（プリンタ登録）, `printer_driver_uninstall.ps1`（ドライバ削除）。`tools/7z.exe` を同梱

## 目的
プリンタドライバとプリンタの登録/削除を行う 3 スクリプト構成のモジュール。
`INF/` 配下に置いた EXE/ZIP 自己解凍アーカイブを 7z.exe で自動展開し、
INF 内のドライバ名と hostlist + `printer_driver_list.csv` の宣言を完全一致照合して
自動インストールする「Auto mode」と、対話選択する「Interactive mode」を持ちます。
プリンタ登録は hostlist 環境変数（最大 10 台）と `printer_list.csv` の union 適用で
10 台超や共通プリンタ宣言にも対応します。

## 入力 (CSV)
**`printer_driver_list.csv`（任意・ドライバ拡張）**:
- `Enabled`, `TargetHost`（空=全ホスト、値=NewPCName 一致）, `DriverName`（INF モデル名と完全一致）, `Description`

**`printer_list.csv`（任意・プリンタ拡張、`printer_delete` がクロス参照）**:
- `Enabled`, `TargetHost`, `PrinterName`, `DriverName`, `PortAddress`, `Description`

hostlist 環境変数: `SELECTED_PRINTER_1..10_NAME / DRIVER / PORT`, `SELECTED_NEW_PCNAME`

## 主要ステップ（Printer Drivers / install）
0. `INF/` 直下の .exe / .zip を `tools/7z.exe` で同名フォルダに自動展開（冪等）
   - 7z 不在時は .zip のみ `Expand-Archive` でフォールバック
1. hostlist + `printer_driver_list.csv` から要求ドライバ名を union 収集
2. `INF/` 配下の全 INF をスキャンし `[Manufacturer]` + アーキテクチャモデルセクションから
   ドライバ名を抽出 → driverMap 構築
3. 要求名 vs INF 内ドライバ名を完全一致照合（[マッチ] / [Unmatched]）
4. 実行確認（Auto mode の AutoPilot は自動 Y）
5. `pnputil /add-driver <inf> /install` で Driver Store 登録 → DriverStore の oem*.inf 経路を解決
6. `Add-PrinterDriver -Name <model> -InfPath <store>` で登録
7. Step 5.5: `Get-PrinterDriver -Name` で各ドライバ存在確認

## 主要ステップ（Register Printers）
1. hostlist 環境変数 + `printer_list.csv` から printers 配列を union
2. ドライバ存在確認（`Get-PrinterDriver`）→ 不足は警告
3. 確認 → `Add-PrinterPort -PrinterHostAddress <IP>` でポート IP_<addr> 作成
4. `Add-Printer -Name -DriverName -PortName` で登録（既存はスキップ）
5. Step 5.5: `Get-Printer` / `Get-PrinterPort` で 4 項目（存在 / DriverName / PortName /
   PrinterHostAddress）を完全一致検証

## 主要ステップ（Uninstall Drivers）
1. `Get-PrinterDriver` 列挙 → Microsoft 標準を除外したリスト表示
2. 番号入力で対象選択
3. 該当ドライバを使うプリンタを事前削除
4. `Restart-Service spooler` でロック解放
5. `Remove-PrinterDriver` 実行 → `pnputil /enum-drivers` から oemN.inf を解決し
   `pnputil /delete-driver oemN.inf /force` で Driver Store からも削除

## 注意点・運用メモ
- 管理者権限必須（pnputil / Add-PrinterDriver / Add-Printer / Spooler restart 全て）
- `INF/` 直下の EXE/ZIP は冪等展開。再展開は手動でフォルダ削除が必要
- `tools/7z.exe` + `7z.dll` は GNU LGPL v2.1+ で同梱、`README-license.txt` で
  バージョン明示。`THIRD_PARTY_NOTICES.md` も連携
- INF 解析はアーキテクチャ（NTamd64 / NTx86）を `[Environment]::Is64BitOperatingSystem` で判定
- 仮想プリンタ（Microsoft Print to PDF 等）は Microsoft 標準パターンで Uninstall 対象から除外

## 検証
- install: `Get-PrinterDriver -Name <model>` で各 matched driver の Driver Store 登録を検証、
  全件存在で `-Verified $true`
- register: 各プリンタに対し Printer 名 / DriverName / PortName /
  PrinterHostAddress の 4 項目完全一致を検証、1 件でも不一致で `$false`
  （reason: not found / driver mismatch / port mismatch / binding mismatch を個別表示）
- uninstall: Verification 未実装、結果サマリのみ
