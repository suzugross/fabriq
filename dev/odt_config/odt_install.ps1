<#
.SYNOPSIS
    Office Deployment Tool (ODT) スタンドアロンインストーラー (修正版)
    
.DESCRIPTION
    configuration.xml 内の <Add> 要素にある SourcePath 属性を
    現在のフォルダの絶対パスに書き換えてインストールを実行します。
    
.PARAMETER XmlFileName
    同じフォルダにある構成ファイルのファイル名（デフォルト: configuration.xml）
#>
Param(
    [string]$XmlFileName = "configuration.xml"
)

# 簡易ログ出力関数
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $TimeStamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $Color = switch ($Level) {
        "INFO" { "Cyan" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        Default { "White" }
    }
    Write-Host "[$TimeStamp][$Level] $Message" -ForegroundColor $Color
}

$ErrorActionPreference = "Stop"

try {
    Write-Log "処理を開始します。"

    # 1. パスの設定（スクリプトの配置場所を基準にする）
    $BaseDir = $PSScriptRoot
    $SetupExePath = Join-Path $BaseDir "setup.exe"
    $ConfigXmlPath = Join-Path $BaseDir $XmlFileName
    $TempXmlPath = Join-Path $BaseDir "temp_config_$(Get-Date -Format 'yyyyMMddHHmmss').xml"

    Write-Log "作業ディレクトリ: $BaseDir"

    # 2. 資材チェック
    if (-not (Test-Path $SetupExePath)) {
        throw "setup.exe が見つかりません。配置場所: $SetupExePath"
    }
    if (-not (Test-Path $ConfigXmlPath)) {
        throw "構成ファイルが見つかりません。配置場所: $ConfigXmlPath"
    }

    # 3. XMLの動的書き換え (Addタグへの絶対パス注入)
    Write-Log "構成ファイル($XmlFileName)を読み込み、<Add>タグの SourcePath を設定します..."
    
    # XML読込
    $XmlContent = [xml](Get-Content $ConfigXmlPath -Encoding UTF8)
    
    # <Configuration> ノードの確認
    if ($null -eq $XmlContent.Configuration) {
        throw "XMLファイルに <Configuration> ノードが見つかりません。"
    }

    # <Add> ノードの取得
    $AddNode = $XmlContent.Configuration.Add
    if ($null -eq $AddNode) {
        throw "XMLファイルに <Add> ノードが見つかりません。構成ファイルを確認してください。"
    }

    # SourcePath 属性を現在の絶対パスで上書き設定
    # 元が "C:\" などになっていても、ここで $BaseDir に書き換わります
    $AddNode.SetAttribute("SourcePath", $BaseDir)
    
    # 一時ファイルとして保存
    $XmlContent.Save($TempXmlPath)
    Write-Log "一時構成ファイルを作成しました: $TempXmlPath"
    Write-Log "適用されたSourcePath: $BaseDir"

    # 4. インストール実行
    Write-Log "Officeのインストールを開始します。完了まで待機してください..."
    
    # 引数: /configure "一時ファイルの絶対パス"
    $Arguments = "/configure `"$TempXmlPath`""
    
    # プロセス起動 (Waitで待機)
    $Process = Start-Process -FilePath $SetupExePath -ArgumentList $Arguments -Wait -NoNewWindow -PassThru
    
    # 5. 結果確認
    if ($Process.ExitCode -eq 0) {
        Write-Log "インストールが正常に完了しました。(ExitCode: 0)" "INFO"
    } else {
        Write-Log "インストールが 警告 または エラー で終了しました。(ExitCode: $($Process.ExitCode))" "WARNING"
        Write-Log "詳細は C:\Windows\Temp などのODTログを確認してください。" "WARNING"
    }

} catch {
    Write-Log "エラーが発生しました: $($_.Exception.Message)" "ERROR"
    if ($Host.Name -eq "ConsoleHost") {
        Read-Host "Enterキーを押して終了してください..."
    }
    exit 1
} finally {
    # 6. 後始末
    if (Test-Path $TempXmlPath) {
        try {
            Remove-Item -Force $TempXmlPath
            Write-Log "一時構成ファイルを削除しました。"
        } catch {
            Write-Log "一時構成ファイルの削除に失敗しました: $TempXmlPath" "WARNING"
        }
    }
    Write-Log "全処理が終了しました。"
}