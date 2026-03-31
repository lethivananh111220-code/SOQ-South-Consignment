Add-Type -AssemblyName System.IO.Compression.FileSystem
$f = "D:\DHF\CONSIGNMENT\Don giao 0104\Lịch 2303-0504.xlsx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($f)
$sstEntry = $zip.GetEntry('xl/sharedStrings.xml')
$strings = @()
if ($sstEntry) {
    $xml = [xml](New-Object System.IO.StreamReader($sstEntry.Open())).ReadToEnd()
    $strings = $xml.sst.si | ForEach-Object { if ($_.t) { $_.t } else { $_.InnerText } }
}
$wsEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
if (-not $wsEntry) { $wsEntry = $zip.GetEntry('xl/worksheets/Sheet1.xml') }
if ($wsEntry) {
    $xml = [xml](New-Object System.IO.StreamReader($wsEntry.Open())).ReadToEnd()
    $row = $xml.worksheet.sheetData.row | Select-Object -First 1
    $headers = @()
    foreach ($c in $row.c) {
        if ($c.t -eq 's') { 
            $val = $strings[[int]$c.v]
            if ($val -is [System.Xml.XmlElement]) { $val = $val.InnerText }
            $headers += $val
        } else {
            $headers += $c.v
        }
    }
    Write-Output ($headers -join '|')
}
$zip.Dispose()
