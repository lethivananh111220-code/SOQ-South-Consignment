$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$wb = $xl.Workbooks.Open("D:\DHF\CONSIGNMENT\Don giao 0204\W13 DHF 2303-2903.xlsx")
$ws = $wb.Sheets.Item(1)
Write-Host "WEEKLY FILE ROW 1:"
for ($j = 1; $j -le 10; $j++) {
    Write-Host -NoNewline ($ws.Cells.Item(1, $j).Text + " | ")
}
Write-Host "`nWEEKLY FILE ROW 2:"
for ($j = 1; $j -le 10; $j++) {
    Write-Host -NoNewline ($ws.Cells.Item(2, $j).Text + " | ")
}
Write-Host "`nWEEKLY FILE ROW 3:"
for ($j = 1; $j -le 10; $j++) {
    Write-Host -NoNewline ($ws.Cells.Item(3, $j).Text + " | ")
}
Write-Host "`n"
$wb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
