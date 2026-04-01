$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$wb = $xl.Workbooks.Open("D:\DHF\CONSIGNMENT\Don giao 0204\W13 DHF 2303-2903.xlsx")
$ws = $wb.Sheets.Item(1)
$headerRow = ""
for ($j = 1; $j -le 15; $j++) {
    $headerRow += $ws.Cells.Item(1, $j).Text + " | "
}
Write-Host "WEEKLY FILE ROW 1: $headerRow"
$headerRow2 = ""
for ($j = 1; $j -le 15; $j++) {
    $headerRow2 += $ws.Cells.Item(2, $j).Text + " | "
}
Write-Host "WEEKLY FILE ROW 2: $headerRow2"
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
