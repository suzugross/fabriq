# wallpaper_config (Standard)

**カテゴリ**: Desktop
**メニュー名**: Wallpaper Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（API 戻り値判定のみ。レジストリ読み返しなし、`-Verified` 未渡し）
**サブスクリプト**: なし（`SystemParametersInfo` を `Add-Type` で同居）

## 目的
デスクトップ壁紙を画像または単色 (SolidColor) で設定するモジュール。
Windows API (`SystemParametersInfo SPI_SETDESKWALLPAPER`) で即時反映するため再起動不要。
画像壁紙では Style（Fill / Fit / Stretch / Tile / Center / Span）を指定でき、
SolidColor では RGB 値で単色背景を直接適用する。`Resolve-HkcuRoot` 経由で SYSTEM 起動時にも
ログオンユーザーの HKCU ハイブにリダイレクトして書き込む昇格対応。

## 入力 (CSV)
`wallpaper_list.csv`
- `Enabled`: 1=実行 / 0=スキップ
- `Type`: `Image` / `SolidColor`（省略時 Image）
- `FileName`: 画像ファイル名（`wallpaper/` 配下相対パス）または絶対パス（Type=Image 時）
- `Style`: `Fill` / `Fit` / `Stretch` / `Tile` / `Center` / `Span`（省略時 Fill）
- `Color`: RGB 値スペース区切り（例 `0 99 177`、Type=SolidColor 時）
- `Description`: 説明
- `Segment`: Segment フィルタ（任意）

サンプル: green.jpg (Image, Fill) / Navy Blue (SolidColor, RGB 0 99 177) いずれも Enabled=0。
画像対応形式: jpg, jpeg, png, bmp, gif, tif, tiff。

## 主要ステップ
1. `wallpaper_list.csv` 読み込み
2. 画像ファイル存在 + 形式確認、または RGB 値妥当性確認 → 一覧表示
3. 実行確認（AutoPilot 自動 Y）
4. Type 別適用:
   - **Image**: `SystemParametersInfo(SPI_SETDESKWALLPAPER, ..., path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)`
     + `Control Panel\Desktop\WallpaperStyle` / `TileWallpaper` レジストリ書き込み
   - **SolidColor**: 壁紙パスをクリア + `Control Panel\Colors\Background` に
     `R G B` 文字列を書き込み + `SystemParametersInfo` トリガー
5. `New-BatchResult` 返却（`-Verified` 未渡し）

## 注意点・運用メモ
- 管理者権限必須（`Resolve-HkcuRoot` で別ユーザーハイブに書き込む可能性のため）
- `wallpaper/` ディレクトリは module ルート配下で、画像をここに置けば相対パス指定可能。
  絶対パスでも OK
- Style 省略時のデフォルトは Fill
- SolidColor 時は FileName / Style は無視
- Sysprep 配布パターンでは Default プロファイルへのコピーは別途必要（本モジュールは
  カレントセッションのみ対象）

## 検証
Post-Apply Verification は実装されていない。
壁紙適用は `SystemParametersInfo` の戻り値（成功 = 非ゼロ）で成否を判定するが、
適用後にレジストリ (`Control Panel\Desktop\WallPaper` / `WallpaperStyle` /
`Colors\Background` 等) を読み返して期待値と一致するかを検証する Step 5.5 はなし。
`-Verified` を `New-BatchResult` に渡していないため履歴 Verified 列は空欄。

理由としては、即時反映 API の戻り値が「OS が壁紙設定を受理した」ことを示すため通常は十分で、
かつレジストリ値（特に Style）の表現が Windows バージョンによって微妙に異なる
（Fill=10/PosV=...）ため一意な比較ロジックを書きにくい、という実装判断。
事実上の確認はオペレータの目視 or `psr.exe` でのスクリーンキャプチャ。
