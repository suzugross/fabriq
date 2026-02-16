# ========================================
# Function: Show Manifesto GUI
# ========================================
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

function Show-Manifesto {
    # Borderless form with paper-white theme
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Manifeste du Surkitinisme"
    $form.Size = New-Object System.Drawing.Size(900, 680)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::FromArgb(250, 248, 244)
    $form.FormBorderStyle = "None"

    # Drop shadow border (thin gray line around borderless form)
    $form.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 200, 200), 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width - 1, $s.Height - 1)
        $pen.Dispose()
    })

    # Font setup
    $fontTitle = New-Object System.Drawing.Font("Meiryo UI", 18, [System.Drawing.FontStyle]::Bold)
    $fontSub   = New-Object System.Drawing.Font("Meiryo UI", 12, [System.Drawing.FontStyle]::Italic)
    $fontBody  = New-Object System.Drawing.Font("Meiryo UI", 11, [System.Drawing.FontStyle]::Regular)

    # Window drag logic (WM_NCLBUTTONDOWN)
    $dragAction = {
        if ($_.Button -eq 'Left') {
            $form.Capture = $false
            $msg = [System.Windows.Forms.Message]::Create($form.Handle, 0xA1, [IntPtr]2, [IntPtr]0)
            $form.DefWndProc([ref]$msg)
        }
    }

    # --- Header area ---
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = "Top"
    $pnlHeader.Height = 100
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 65)
    $pnlHeader.Add_MouseDown($dragAction)
    $form.Controls.Add($pnlHeader)

    # Title label
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Manifeste du Surkitinisme"
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 220, 220)
    $lblTitle.Font = $fontTitle
    $lblTitle.AutoSize = $false
    $lblTitle.TextAlign = "MiddleCenter"
    $lblTitle.Size = New-Object System.Drawing.Size(900, 50)
    $lblTitle.Location = New-Object System.Drawing.Point(0, 15)
    $lblTitle.BackColor = [System.Drawing.Color]::Transparent
    $lblTitle.Add_MouseDown($dragAction)
    $pnlHeader.Controls.Add($lblTitle)

    # Subtitle
    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text = [char]0x2015 + [char]0x2015 + " " + [char]0x30B7 + [char]0x30E5 + [char]0x30EB + [char]0x30AD + [char]0x30C6 + [char]0x30A3 + [char]0x30CB + [char]0x30B9 + [char]0x30E0 + [char]0x5BA3 + [char]0x8A00 + " " + [char]0x2015 + [char]0x2015
    $lblSub.ForeColor = [System.Drawing.Color]::FromArgb(190, 190, 190)
    $lblSub.Font = $fontSub
    $lblSub.AutoSize = $false
    $lblSub.TextAlign = "MiddleCenter"
    $lblSub.Size = New-Object System.Drawing.Size(900, 30)
    $lblSub.Location = New-Object System.Drawing.Point(0, 60)
    $lblSub.BackColor = [System.Drawing.Color]::Transparent
    $lblSub.Add_MouseDown($dragAction)
    $pnlHeader.Controls.Add($lblSub)

    # --- Body area ---
    $manifestoText = @"
キッティングエンジニアとして生きるということに対する、その構築という上での最も不確実な部分、つまり、いうまでもなく、その「手動によるOS設定」に対する信仰が高じすぎると、最後には、その信仰は失われてしまう。キッティングという名のこの決定的な夢想家は、日に日に自分の環境構築への不満を募らせ、仕方なく叩く羽目に至ってきたコマンド群を、苦労して調べまわしてみるのである。

そういうパラメータ群は、無頓着なマニュアル更新によって、あるいは「根性」という名の努力によって、いや、たいていはこの不毛な努力によってもたらされたものである。というのは、彼は泥臭い作業に同意したからであり、少なくとも「現場の運」（運と称しているものを！）を賭けることをいとわなかったからである。

そうなると、エンジニアの得られる分け前はとてもつつましいものである。どんなレガシーなドライバを掴まされたか、どんなくだらぬ不具合修正に足を突っ込んできたか、これは各自みなご承知の通りである。スキルの有無も全く関係なく、この点においては誰もみな「工場出荷状態のPC」と変わりなく、それにまた標準化意識を口にしてみたところで、そんなものなどなくても人間は平気で（ただ漫然と）作業を続けられることを認めよう。

いくらかでも効率化の正気をとどめていれば、このとき、エンジニアは自分の「最初のコード（プロトタイプ）」を頼りにするしかない。これだけは、保守管理者たちのおせっかいのせいで、どれほどスパゲッティコードにされていたとしても、彼の目にはやはり魅力溢れたオートメーションとして映るからだ。
そこでは、不整合として知られているものが一切存在しないため、同時にいくつもの展開プロファイルの展望を許される。エンジニアはその「完全自動化」という幻想の中に根を下ろし、もはやあらゆる設定の、つかの間の、極端な容易さしか認めようとしない。

毎朝、キッティングエンジニアたちは不安なしに Fabriq.bat を叩く。すべては hostlist.csv にあり、最悪のネットワーク条件でさえも素晴らしい。進捗ステータスは緑にもなれば赤にもなる。
プロセスは決して眠らないだろう。
"@

    # RichTextBox for body text (scrollbar pushed to far right, away from text)
    $txtContent = New-Object System.Windows.Forms.RichTextBox
    $txtContent.Text = $manifestoText
    $txtContent.Font = $fontBody
    $txtContent.ForeColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $txtContent.BackColor = [System.Drawing.Color]::FromArgb(250, 248, 244)
    $txtContent.BorderStyle = "None"
    $txtContent.ReadOnly = $true
    $txtContent.ScrollBars = "Vertical"
    $txtContent.Location = New-Object System.Drawing.Point(80, 130)
    $txtContent.Size = New-Object System.Drawing.Size(800, 450)
    $txtContent.RightMargin = 680
    $txtContent.Cursor = [System.Windows.Forms.Cursors]::Default
    $form.Controls.Add($txtContent)

    # --- Footer area ---
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Font = New-Object System.Drawing.Font("Meiryo UI", 10)
    $btnClose.FlatStyle = "Flat"
    $btnClose.FlatAppearance.BorderSize = 1
    $btnClose.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
    $btnClose.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $btnClose.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(240, 238, 234)
    $btnClose.Size = New-Object System.Drawing.Size(200, 45)
    $btnClose.Location = New-Object System.Drawing.Point(350, 600)
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() })
    $form.Controls.Add($btnClose)

    # Keyboard shortcut (ESC to close)
    $form.KeyPreview = $true
    $form.Add_KeyDown({
        if ($_.KeyCode -eq "Escape") { $form.Close() }
    })

    # Show form
    $form.Add_Shown({ $form.Activate(); $btnClose.Focus() })
    [void]$form.ShowDialog()
    $form.Dispose()
}
