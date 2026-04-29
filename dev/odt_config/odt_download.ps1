<#
.SYNOPSIS
    Office Deployment Tool (ODT) standalone installer (revised).

.DESCRIPTION
    Rewrites the SourcePath attribute on the <Add> element inside
    configuration.xml to the current folder's absolute path, then
    runs the installer.

.PARAMETER XmlFileName
    Configuration filename in the same folder (default: configuration.xml).
#>
Param(
    [string]$XmlFileName = "configuration.xml"
)

# Lightweight logging helper
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

    # 1. Path setup (anchored to the script location)
    $BaseDir = $PSScriptRoot
    $SetupExePath = Join-Path $BaseDir "setup.exe"
    $ConfigXmlPath = Join-Path $BaseDir $XmlFileName
    $TempXmlPath = Join-Path $BaseDir "temp_config_$(Get-Date -Format 'yyyyMMddHHmmss').xml"

    Write-Log "作業ディレクトリ: $BaseDir"

    # 2. Material check
    if (-not (Test-Path $SetupExePath)) {
        throw "setup.exe が見つかりません。配置場所: $SetupExePath"
    }
    if (-not (Test-Path $ConfigXmlPath)) {
        throw "構成ファイルが見つかりません。配置場所: $ConfigXmlPath"
    }

    # 3. Dynamic XML rewrite (inject absolute path into the <Add> tag)
    Write-Log "構成ファイル($XmlFileName)を読み込み、<Add>タグの SourcePath を設定します..."
    
    # Load the XML
    $XmlContent = [xml](Get-Content $ConfigXmlPath -Encoding UTF8)
    
    # Verify the <Configuration> node exists
    if ($null -eq $XmlContent.Configuration) {
        throw "XMLファイルに <Configuration> ノードが見つかりません。"
    }

    # Get the <Add> node
    $AddNode = $XmlContent.Configuration.Add
    if ($null -eq $AddNode) {
        throw "XMLファイルに <Add> ノードが見つかりません。構成ファイルを確認してください。"
    }

    # Overwrite the SourcePath attribute with the current absolute path.
    # Any prior value (e.g., "C:\") is replaced with $BaseDir here.
    $AddNode.SetAttribute("SourcePath", $BaseDir)
    
    # Save as a temp file
    $XmlContent.Save($TempXmlPath)
    Write-Log "一時構成ファイルを作成しました: $TempXmlPath"
    Write-Log "適用されたSourcePath: $BaseDir"

    # 4. Run the installer
    Write-Log "Officeのインストールを開始します。完了まで待機してください..."
    
    # Args: /configure "<absolute path to temp file>"
    $Arguments = "/configure `"$TempXmlPath`""
    
    # Start the process (block via -Wait)
    $Process = Start-Process -FilePath $SetupExePath -ArgumentList $Arguments -Wait -NoNewWindow -PassThru
    
    # 5. Check the result
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
    # 6. Cleanup
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