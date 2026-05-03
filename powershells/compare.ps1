param([string]$SourceFolder, [string]$TargetFolder)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$totalPeopleCount = 0

# Excelプロセスを一度クリア
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

$forbiddenFile = Join-Path $PSScriptRoot "禁止ワード.txt"
$forbiddenWords = @()
if (Test-Path $forbiddenFile) {
    $forbiddenWords = Get-Content $forbiddenFile -Encoding Default | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# エラーの原因となる Calculation の変更を削除し、ScreenUpdating のみに制限
$excel.ScreenUpdating = $false

$config = @{
    TargetPrefix = "日報"; StartCol = "H"; DiffStatusCol = "J"; ForbiddenCol = "K"
    WorkColName = "作業内容"; LabelCols = @("出勤", "退勤")
}
$compareMap = @{ "開始時間" = "出勤"; "終了時間" = "退勤" }

function Get-Details($sheet) {
    # 最初の100行40列を一括取得してスキャン
    $range = $sheet.Range("A1:AN100").Value2
    if ($null -eq $range) { return $null }
    for ($r = 1; $r -le 100; $r++) {
        $map = @{}
        for ($c = 1; $c -le 40; $c++) {
            $txt = [string]$range[$r, $c]
            if (![string]::IsNullOrWhiteSpace($txt)) { $map[$txt.Trim()] = $c }
        }
        if ($map.ContainsKey("日付")) { return [PSCustomObject]@{ Row = $r; Map = $map } }
    }
    return $null
}

try {
    if ([string]::IsNullOrWhiteSpace($SourceFolder)) { throw "フォルダパスが空です。" }

    $files = Get-ChildItem $SourceFolder, $TargetFolder -Filter *.xlsx -ErrorAction SilentlyContinue
    $people = @{}
    foreach ($f in $files) {
        $parts = ([System.IO.Path]::GetFileNameWithoutExtension($f.Name)).Split("_")
        if ($parts.Count -lt 2) { continue }
        $person = $parts[-1]; $prefix = $parts[0]
        if (-not $people.ContainsKey($person)) { $people[$person] = @{} }
        $people[$person][$prefix] = $f.FullName
    }

    $totalPeopleCount = $people.Count
    "PROGRESS:TOTAL:$totalPeopleCount"

    foreach ($person in $people.Keys) {
        "PROGRESS:CURRENT:$person 処理中..."
        $pair = $people[$person]
        if ($pair.Count -lt 2) { continue }

        $book1 = $excel.Workbooks.Open($pair[($pair.Keys | Select-Object -First 1)], 0, $false)
        $book2 = $excel.Workbooks.Open($pair[($pair.Keys | Select-Object -Skip 1 -First 1)], 0, $false)
        $s1 = $book1.Sheets.Item(1); $s2 = $book2.Sheets.Item(1)
        
        $d1 = Get-Details $s1; $d2 = Get-Details $s2
        if ($null -eq $d1 -or $null -eq $d2) { $book1.Close($false); $book2.Close($false); continue }

        $isS1Target = ($pair.Keys | Select-Object -First 1) -like "*$($config.TargetPrefix)*"
        $tSheet = if ($isS1Target) { $s1 } else { $s2 }; $tDet = if ($isS1Target) { $d1 } else { $d2 }
        $oSheet = if ($isS1Target) { $s2 } else { $s1 }; $oDet = if ($isS1Target) { $d2 } else { $d1 }

        # データ範囲を一括取得
        $maxR = 1000 
        $tData = $tSheet.Range($tSheet.Cells.Item(1, 1), $tSheet.Cells.Item($maxR, 40)).Value2
        $oData = $oSheet.Range($oSheet.Cells.Item(1, 1), $oSheet.Cells.Item($maxR, 40)).Value2

        # 勤怠側を辞書化
        $oDict = @{}
        for ($r = $oDet.Row + 1; $r -le $maxR; $r++) {
            $d = [string]$oData[$r, $oDet.Map["日付"]]
            if ([string]::IsNullOrWhiteSpace($d)) { break }
            $oDict[$d] = $r
        }

        # 見出し書き込み
        $hRow = $tDet.Row
        $startColIdx = $tSheet.Range($config.StartCol + "1").Column
        $diffColIdx = $tSheet.Range($config.DiffStatusCol + "1").Column
        $fbdColIdx = $tSheet.Range($config.ForbiddenCol + "1").Column

        $tSheet.Cells.Item($hRow, $startColIdx).Value2 = "勤怠出勤"
        $tSheet.Cells.Item($hRow, $startColIdx + 1).Value2 = "勤怠退勤"
        $tSheet.Cells.Item($hRow, $diffColIdx).Value2 = "差分判定"
        $tSheet.Cells.Item($hRow, $fbdColIdx).Value2 = "禁則チェック"

        $diffCount = 0; $fbdCount = 0
        
        for ($rT = $tDet.Row + 1; $rT -le $maxR; $rT++) {
            $dateT = [string]$tData[$rT, $tDet.Map["日付"]]
            if ([string]::IsNullOrWhiteSpace($dateT)) { break }

            # 禁則チェック
            $checkResult = ""
            if ($tDet.Map.ContainsKey($config.WorkColName)) {
                $workContent = [string]$tData[$rT, $tDet.Map[$config.WorkColName]]
                foreach ($word in $forbiddenWords) {
                    if ($workContent.Contains($word)) { $checkResult = "禁則あり($word)"; $fbdCount++; break }
                }
            }
            if ($checkResult -ne "") {
                $tSheet.Cells.Item($rT, $fbdColIdx).Value2 = $checkResult
                $tSheet.Cells.Item($rT, $fbdColIdx).Font.Color = 255
            }

            # 日付比較
            if ($oDict.ContainsKey($dateT)) {
                $rO = $oDict[$dateT]
                $isDiff = $false
                foreach ($kT in $compareMap.Keys) {
                    $kO = $compareMap[$kT]
                    if ($tDet.Map.ContainsKey($kT) -and $oDet.Map.ContainsKey($kO)) {
                        $vT = ([string]$tData[$rT, $tDet.Map[$kT]]).Trim()
                        $vO = ([string]$oData[$rO, $oDet.Map[$kO]]).Trim()
                        if ($vT -ne $vO) { $isDiff = $true }
                    }
                }

                if ($isDiff) {
                    $diffCount++
                    $diffCell = $tSheet.Cells.Item($rT, $diffColIdx)
                    $diffCell.Value2 = "差分あり"
                    $diffCell.Font.Color = 255
                }

                # 時間転記
                for($i=0; $i -lt $config.LabelCols.Count; $i++) {
                    $label = $config.LabelCols[$i]
                    if ($oDet.Map.ContainsKey($label)) {
                        $val = [string]$oData[$rO, $oDet.Map[$label]]
                        if (![string]::IsNullOrWhiteSpace($val)) {
                            $cell = $tSheet.Cells.Item($rT, $startColIdx + $i)
                            $cell.NumberFormatLocal = "h:mm"
                            $cell.Value2 = $val
                            if ($isDiff) { $cell.Font.Color = 255 }
                        }
                    }
                }
            }
        }
        $book1.Save(); $book2.Save(); $book1.Close(); $book2.Close()
        "$person : 差分 $diffCount 件 / 禁則 $fbdCount 件"
    }
}
catch { "ERROR: $($_.Exception.Message)" }
finally {
    $sw.Stop()
    $time = "$($sw.Elapsed.Minutes)分$($sw.Elapsed.Seconds)秒"
    "RESULT:処理件数: $totalPeopleCount 件 / 処理時間: $time"
    if ($excel) { 
        $excel.ScreenUpdating = $true
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null 
    }
}