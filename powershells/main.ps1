Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===============================
# デフォルト設定
# ===============================
$defaultFolder1 = "C:\Users\masaki\Desktop\sample\日報"
$defaultFolder2 = "C:\Users\masaki\Desktop\sample\勤怠"

# ===============================
# 共通実行関数
# ===============================
function Run-ExternalScript {
    param(
        [string]$FileName,
        [string]$Arguments,
        [string]$FinishTitle
    )

    $listBox.Items.Clear()
    $startButton.Enabled = $false
    $clearButton.Enabled = $false
    
    $scriptPath = Join-Path $PSScriptRoot $FileName
    
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" $Arguments"
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8 

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $finalResult = "処理が完了しました"

    if ($process.Start()) {
        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            if ($null -ne $line -and ![string]::IsNullOrWhiteSpace($line)) {
                
                if ($line -like "PROGRESS:*") {
                    $msg = $line -replace "PROGRESS:", ""
                    if ($listBox.Items.Count -eq 0) {
                        $null = $listBox.Items.Add($msg)
                    } else {
                        $listBox.Items[0] = $msg
                    }
                } elseif ($line -like "RESULT:*") {
                    $finalResult = $line -replace "RESULT:", ""
                    $null = $listBox.Items.Add($line)
                } else {
                    # --- フィルタリングロジックの強化 ---
                    # 1. まず「差分」と「禁則」の数値部分を抽出
                    if ($line -match "差分\s+(\d+)\s+件\s+/\s+禁則\s+(\d+)\s+件") {
                        $diffCount = [int]$matches[1]
                        $fbdCount = [int]$matches[2]

                        # どちらかが1件以上の場合のみリストに追加
                        if ($diffCount -gt 0 -or $fbdCount -gt 0) {
                            $null = $listBox.Items.Add($line)
                        }
                    } 
                    # 2. それ以外のエラーメッセージなどは念のため表示
                    elseif ($line -like "ERROR:*") {
                        $null = $listBox.Items.Add($line)
                    }
                }
                
                $listBox.TopIndex = $listBox.Items.Count - 1
                $listBox.Refresh()
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        $process.WaitForExit()
    }

    $process.Close()
    $startButton.Enabled = $true
    $clearButton.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show($finalResult, $FinishTitle)
}

# ===============================
# フォーム構成
# ===============================
$form = New-Object Windows.Forms.Form
$form.Text = "フォルダ比較ツール"
$form.Size = New-Object Drawing.Size(650, 450)
$form.StartPosition = "CenterScreen"

$textBox1 = New-Object Windows.Forms.TextBox
$textBox1.Location = New-Object Drawing.Point(10, 10); $textBox1.Size = New-Object Drawing.Size(450, 25); $textBox1.Text = $defaultFolder1
$form.Controls.Add($textBox1)

$btnSelect1 = New-Object Windows.Forms.Button
$btnSelect1.Text = "作業対象フォルダ選択"; $btnSelect1.Location = New-Object Drawing.Point(470, 10); $btnSelect1.Size = New-Object Drawing.Size(150, 25)
$btnSelect1.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $textBox1.Text
    if ($dialog.ShowDialog() -eq "OK") { $textBox1.Text = $dialog.SelectedPath }
})
$form.Controls.Add($btnSelect1)

$textBox2 = New-Object Windows.Forms.TextBox
$textBox2.Location = New-Object Drawing.Point(10, 50); $textBox2.Size = New-Object Drawing.Size(450, 25); $textBox2.Text = $defaultFolder2
$form.Controls.Add($textBox2)

$btnSelect2 = New-Object Windows.Forms.Button
$btnSelect2.Text = "比較対象フォルダ選択"; $btnSelect2.Location = New-Object Drawing.Point(470, 50); $btnSelect2.Size = New-Object Drawing.Size(150, 25)
$btnSelect2.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $textBox2.Text
    if ($dialog.ShowDialog() -eq "OK") { $textBox2.Text = $dialog.SelectedPath }
})
$form.Controls.Add($btnSelect2)

$startButton = New-Object Windows.Forms.Button
$startButton.Text = "比較開始"; $startButton.Location = New-Object Drawing.Point(10, 90); $startButton.Size = New-Object Drawing.Size(120, 30)
$form.Controls.Add($startButton)

$clearButton = New-Object Windows.Forms.Button
$clearButton.Text = "チェック内容クリア"; $clearButton.Location = New-Object Drawing.Point(200, 90); $clearButton.Size = New-Object Drawing.Size(150, 30)
$form.Controls.Add($clearButton)

$listBox = New-Object Windows.Forms.ListBox
$listBox.Location = New-Object Drawing.Point(10, 140); $listBox.Size = New-Object Drawing.Size(610, 250)
$form.Controls.Add($listBox)

# ===============================
# ダブルクリックでファイルを開く
# ===============================
$listBox.Add_MouseDoubleClick({
    if ($listBox.SelectedItem -eq $null) { return }
    $selectedLine = $listBox.SelectedItem.ToString()
    if ($selectedLine -match "^(.+?)\s*:") {
        $personName = $matches[1].Trim()
        $targetFile = Get-ChildItem -Path $textBox1.Text -Filter "*$personName*.xlsx" | Select-Object -First 1
        if ($targetFile -ne $null) {
            try {
                Start-Process $targetFile.FullName
            } catch {
                [System.Windows.Forms.MessageBox]::Show("ファイルを開けませんでした。")
            }
        }
    }
})

# ===============================
# イベント割り当て
# ===============================
$startButton.Add_Click({
    $args = "-SourceFolder `"$($textBox1.Text)`" -TargetFolder `"$($textBox2.Text)`""
    Run-ExternalScript -FileName "compare.ps1" -Arguments $args -FinishTitle "比較完了"
})

$clearButton.Add_Click({
    $args = "-SourceFolder `"$($textBox1.Text)`""
    Run-ExternalScript -FileName "clear.ps1" -Arguments $args -FinishTitle "クリア完了"
})

$form.ShowDialog()