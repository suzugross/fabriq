Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================================
# グローバル設定 / 初期値
# ==========================================
$Script:DefaultCsvPath = ".\app_list.csv"

# ==========================================
# GUIフォーム作成
# ==========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Winget GUI Editor for Fabriq"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Meiryo UI", 9)

# ------------------------------------------
# レイアウト (SplitContainer)
# ------------------------------------------
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = "Fill"
$splitContainer.Orientation = "Horizontal"
$splitContainer.SplitterDistance = 350
$splitContainer.SplitterWidth = 5
$form.Controls.Add($splitContainer)

# ==========================================
# 上部パネル: Winget 検索エリア
# ==========================================
$panelTop = $splitContainer.Panel1
$panelTop.Padding = New-Object System.Windows.Forms.Padding(10)

# ネット接続状態表示
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10, 10)
$lblStatus.Size = New-Object System.Drawing.Size(300, 20)
$lblStatus.Text = "ネットワーク接続確認中..."
$lblStatus.ForeColor = "Gray"
$panelTop.Controls.Add($lblStatus)

# 検索ボックスとボタン
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "アプリ検索 (キーワード):"
$lblSearch.Location = New-Object System.Drawing.Point(10, 40)
$lblSearch.Size = New-Object System.Drawing.Size(150, 20)
$panelTop.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(160, 38)
$txtSearch.Size = New-Object System.Drawing.Size(300, 25)
$panelTop.Controls.Add($txtSearch)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "検索実行"
$btnSearch.Location = New-Object System.Drawing.Point(470, 36)
$btnSearch.Size = New-Object System.Drawing.Size(100, 28)
$panelTop.Controls.Add($btnSearch)

# 検索結果グリッド
$gridSearch = New-Object System.Windows.Forms.DataGridView
$gridSearch.Location = New-Object System.Drawing.Point(10, 75)
$gridSearch.Size = New-Object System.Drawing.Size(960, 200)
$gridSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$gridSearch.AllowUserToAddRows = $false
$gridSearch.SelectionMode = "FullRowSelect"
$gridSearch.MultiSelect = $false
$gridSearch.ReadOnly = $true
$gridSearch.ColumnCount = 3
$gridSearch.Columns[0].Name = "ID"
$gridSearch.Columns[0].Width = 250
$gridSearch.Columns[1].Name = "Name"
$gridSearch.Columns[1].Width = 300
$gridSearch.Columns[2].Name = "Version"
$gridSearch.Columns[2].Width = 100
$panelTop.Controls.Add($gridSearch)

# 追加オプション入力エリア
$lblOptions = New-Object System.Windows.Forms.Label
$lblOptions.Text = "追加オプション (例: --silent --scope machine):"
$lblOptions.Location = New-Object System.Drawing.Point(10, 290)
$lblOptions.Size = New-Object System.Drawing.Size(300, 20)
$lblOptions.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$panelTop.Controls.Add($lblOptions)

$txtOptions = New-Object System.Windows.Forms.TextBox
$txtOptions.Location = New-Object System.Drawing.Point(10, 310)
$txtOptions.Size = New-Object System.Drawing.Size(400, 25)
$txtOptions.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$panelTop.Controls.Add($txtOptions)

# CSVへ追加ボタン
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "↓↓ CSVリストへ追加 ↓↓"
$btnAdd.Location = New-Object System.Drawing.Point(420, 308)
$btnAdd.Size = New-Object System.Drawing.Size(200, 28)
$btnAdd.BackColor = "LightBlue"
$btnAdd.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$panelTop.Controls.Add($btnAdd)


# ==========================================
# 下部パネル: CSV 編集エリア
# ==========================================
$panelBottom = $splitContainer.Panel2
$panelBottom.Padding = New-Object System.Windows.Forms.Padding(10)

# CSV操作ボタン
$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "CSV読込"
$btnLoad.Location = New-Object System.Drawing.Point(10, 10)
$panelBottom.Controls.Add($btnLoad)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "CSV保存"
$btnSave.Location = New-Object System.Drawing.Point(100, 10)
$panelBottom.Controls.Add($btnSave)

$btnDeleteRow = New-Object System.Windows.Forms.Button
$btnDeleteRow.Text = "選択行を削除"
$btnDeleteRow.Location = New-Object System.Drawing.Point(200, 10)
$btnDeleteRow.Size = New-Object System.Drawing.Size(120, 28)
$panelBottom.Controls.Add($btnDeleteRow)

# CSVグリッド
$gridCsv = New-Object System.Windows.Forms.DataGridView
$gridCsv.Location = New-Object System.Drawing.Point(10, 50)
$gridCsv.Size = New-Object System.Drawing.Size(960, 250)
$gridCsv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$gridCsv.AutoSizeColumnsMode = "Fill"
$panelBottom.Controls.Add($gridCsv)

# データテーブルの準備 (CSV用)
$dt = New-Object System.Data.DataTable
[void]$dt.Columns.Add("Id")
[void]$dt.Columns.Add("Name")
[void]$dt.Columns.Add("Options") # インストールオプション用カラム
$gridCsv.DataSource = $dt

# ==========================================
# イベントハンドラ
# ==========================================

# 1. フォームロード時: Ping確認
$form.Add_Load({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $ping = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
        if ($ping) {
            $lblStatus.Text = "ネットワーク接続: OK (8.8.8.8)"
            $lblStatus.ForeColor = "Green"
        } else {
            $lblStatus.Text = "ネットワーク接続: NG (8.8.8.8 へのPing失敗)"
            $lblStatus.ForeColor = "Red"
        }
    } catch {
        $lblStatus.Text = "ネットワーク接続: エラー"
        $lblStatus.ForeColor = "Red"
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# 2. 検索ボタン: Winget実行
$btnSearch.Add_Click({
    $query = $txtSearch.Text
    if ([string]::IsNullOrWhiteSpace($query)) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $gridSearch.Rows.Clear()
    
    # Wingetプロセス実行 (非同期ではなく簡易的に同期実行)
    try {
        # 文字化け対策: OutputEncodingを指定して実行
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "winget.exe"
        $psi.Arguments = "search `"$query`" --accept-source-agreements"
        $psi.RedirectStandardOutput = $true
        $psi.UseShellExecute = $false
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.CreateNoWindow = $true
        
        $process = [System.Diagnostics.Process]::Start($psi)
        $output = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()

        # 結果のパース (簡易実装: 行ごとに処理)
        $lines = $output -split "`r`n"
        $startParsing = $false
        
        foreach ($line in $lines) {
            if ($line -match "^Name\s+Id\s+Version") {
                $startParsing = $true
                continue
            }
            if (-not $startParsing) { continue }
            if ($line -match "^-+") { continue } # 区切り線
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            # 固定長やスペース区切りの解析は難しいが、IDの特定を試みる
            # 一般的な出力: Name (可変)   Id (識別子)   Version
            # 簡易的に、2文字以上のスペースで分割して取得する
            $parts = $line -split "\s{2,}"
            
            if ($parts.Count -ge 2) {
                $name = $parts[0].Trim()
                $id = $parts[1].Trim()
                $ver = if ($parts.Count -ge 3) { $parts[2].Trim() } else { "" }
                
                # IDっぽくないものは弾く（簡易フィルタ）
                if ($id.Length -gt 2) {
                    $gridSearch.Rows.Add($id, $name, $ver)
                }
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Wingetの実行に失敗しました。`n$_", "エラー", "OK", "Error")
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# 3. 追加ボタン: 検索結果からCSVグリッドへ転記
$btnAdd.Add_Click({
    if ($gridSearch.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("上のリストからアプリを選択してください。", "注意")
        return
    }
    
    $row = $gridSearch.SelectedRows[0]
    $id = $row.Cells["ID"].Value
    $name = $row.Cells["Name"].Value
    $opts = $txtOptions.Text

    # 重複チェック（IDで）
    $existing = $dt.Select("Id = '$id'")
    if ($existing.Count -gt 0) {
        $res = [System.Windows.Forms.MessageBox]::Show("ID: $id は既に追加されています。追加しますか？", "確認", "YesNo", "Question")
        if ($res -eq "No") { return }
    }

    $dt.Rows.Add($id, $name, $opts)
})

# 4. 行削除ボタン
$btnDeleteRow.Add_Click({
    if ($gridCsv.SelectedRows.Count -gt 0) {
        foreach ($row in $gridCsv.SelectedRows) {
            $dt.Rows.RemoveAt($row.Index)
        }
    }
})

# 5. CSV保存
$btnSave.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $sfd.FileName = "app_list.csv"
    
    if ($sfd.ShowDialog() -eq "OK") {
        try {
            # DataTableをCSVとしてエクスポート
            $exportData = @()
            foreach ($row in $dt.Rows) {
                $obj = [PSCustomObject]@{
                    Id      = $row["Id"]
                    Name    = $row["Name"]
                    Options = $row["Options"]
                }
                $exportData += $obj
            }
            $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("保存しました。", "完了")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("保存に失敗しました。`n$_", "エラー")
        }
    }
})

# 6. CSV読込
$btnLoad.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    
    if ($ofd.ShowDialog() -eq "OK") {
        try {
            $csvData = Import-Csv -Path $ofd.FileName -Encoding UTF8
            $dt.Rows.Clear()
            foreach ($item in $csvData) {
                # カラム名が一致する場合のみ読み込む簡易ロジック
                $id = if ($item.PSObject.Properties["Id"]) { $item.Id } else { "" }
                $name = if ($item.PSObject.Properties["Name"]) { $item.Name } else { "" }
                $opt = if ($item.PSObject.Properties["Options"]) { $item.Options } else { "" }
                
                $dt.Rows.Add($id, $name, $opt)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("読み込みに失敗しました。`n$_", "エラー")
        }
    }
})

# ==========================================
# フォーム表示
# ==========================================
[void]$form.ShowDialog()