# ==========================================
# 設定: ここに変更したい解像度を入力してください
# ==========================================
$targetWidth  = 1920  # 横幅 (例: 1920)
$targetHeight = 1080  # 高さ (例: 1080)

# ==========================================
# 内部処理 (C#コードの埋め込み)
# ==========================================
$code = @'
using System;
using System.Runtime.InteropServices;

public class ResolutionHandler {
    [DllImport("user32.dll")]
    public static extern int ChangeDisplaySettings(ref DEVMODE devMode, int flags);

    [StructLayout(LayoutKind.Sequential)]
    public struct DEVMODE {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    // 定数定義
    public const int ENUM_CURRENT_SETTINGS = -1;
    public const int CDS_UPDATEREGISTRY = 0x01;
    public const int CDS_TEST = 0x02;
    public const int DISP_CHANGE_SUCCESSFUL = 0;
    public const int DISP_CHANGE_RESTART = 1;
    public const int DISP_CHANGE_FAILED = -1;

    // フラグ (WidthとHeightを変更することを指定)
    public const int DM_PELSWIDTH = 0x80000;
    public const int DM_PELSHEIGHT = 0x100000;

    public static string ChangeRes(int width, int height) {
        DEVMODE dm = new DEVMODE();
        dm.dmDeviceName = new String(new char[32]);
        dm.dmFormName = new String(new char[32]);
        dm.dmSize = (short)Marshal.SizeOf(dm);

        // 指定された幅と高さをセット
        dm.dmPelsWidth = width;
        dm.dmPelsHeight = height;
        
        // 「幅と高さを変更する」というフラグを立てる
        dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;

        // 設定を適用 (CDS_UPDATEREGISTRYを指定するとレジストリにも保存される)
        int iRet = ChangeDisplaySettings(ref dm, CDS_UPDATEREGISTRY);

        switch (iRet) {
            case DISP_CHANGE_SUCCESSFUL: return "成功: 解像度を変更しました。";
            case DISP_CHANGE_RESTART: return "警告: 反映には再起動が必要です。";
            default: return "失敗: 設定できませんでした。解像度がサポートされていない可能性があります。";
        }
    }
}
'@

# C#コードをコンパイルして読み込む
Add-Type -TypeDefinition $code -Language CSharp

# ==========================================
# 実行部分
# ==========================================
Write-Host "解像度を ${targetWidth} x ${targetHeight} に変更しようとしています..." -ForegroundColor Cyan

try {
    # 定義したクラスのメソッドを呼び出す
    $result = [ResolutionHandler]::ChangeRes($targetWidth, $targetHeight)
    
    if ($result -like "成功*") {
        Write-Host $result -ForegroundColor Green
    } else {
        Write-Host $result -ForegroundColor Red
    }
}
catch {
    Write-Host "エラーが発生しました: $_" -ForegroundColor Red
}

# ウィンドウがすぐ閉じないように一時停止
Write-Host "`n何かキーを押すと終了します..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")