param([string]$SourceFolder)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$total = 0

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$config = @{ TargetPrefix = "“ú•ñ"; DeleteCols = @("K", "J", "I", "H") }

try {
    if (Test-Path $SourceFolder) {
        $files = Get-ChildItem $SourceFolder -Filter "*$($config.TargetPrefix)*.xlsx"
        $total = $files.Count
        "PROGRESS:TOTAL:$total"
        foreach ($f in $files) {
            $person = ([System.IO.Path]::GetFileNameWithoutExtension($f.Name)).Split("_")[-1]
            "PROGRESS:CURRENT:$person ‚Ì—ñ‚ğíœ’†..."
            $book = $excel.Workbooks.Open($f.FullName)
            $sheet = $book.Sheets.Item(1)
            foreach ($colLetter in $config.DeleteCols) {
                $sheet.Columns.Item($colLetter).Delete() | Out-Null
            }
            $book.Save(); $book.Close()
            "$person : —ñíœŠ®—¹"
        }
    }
}
finally {
    $sw.Stop()
    $time = "$($sw.Elapsed.Minutes)•ª$($sw.Elapsed.Seconds)•b"
    "RESULT:ƒNƒŠƒAŒ”: $total Œ / ˆ—ŠÔ: $time"
    $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}