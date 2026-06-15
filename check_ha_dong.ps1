Add-Type -AssemblyName System.IO.Compression.FileSystem

$filePath = Get-ChildItem "..\*1606\merchandiser_report_*.xlsx" | Select-Object -ExpandProperty FullName -First 1
$tempPath = ".\temp_peek_1606_2.zip"

if ($filePath) {
    Copy-Item $filePath $tempPath -Force
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
    $sstEntry = $zip.GetEntry('xl/sharedStrings.xml')
    $strings = @()
    if ($sstEntry) {
        $xml = [xml](New-Object System.IO.StreamReader($sstEntry.Open())).ReadToEnd()
        $strings = $xml.sst.si | ForEach-Object { if ($_.t) { $_.t } else { $_.InnerText } }
    }

    $wsEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
    if (-not $wsEntry) { $wsEntry = $zip.GetEntry('xl/worksheets/Sheet1.xml') }

    $dates = @{}
    if ($wsEntry) {
        $xml = [xml](New-Object System.IO.StreamReader($wsEntry.Open())).ReadToEnd()
        $rows = $xml.worksheet.sheetData.row
        foreach ($r in $rows) {
            $vals = @()
            $match = $false
            foreach ($c in $r.c) {
                $val = ''
                if ($c.t -eq 's') { 
                    $val = $strings[[int]$c.v]
                    if ($val -is [System.Xml.XmlElement]) { $val = $val.InnerText }
                } else {
                    $val = $c.v
                }
                if ($val -match "1530") { $match = $true }
                $vals += $val
            }
            if ($match) {
                $dateVal = $vals[2]
                if (-not $dates.ContainsKey($dateVal)) {
                    $dates[$dateVal] = 1
                } else {
                    $dates[$dateVal]++
                }
            }
        }
        $dates.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Output "$($_.Name): $($_.Value) products" }
    }
    $zip.Dispose()
    Remove-Item $tempPath -Force
} else {
    Write-Output "File not found"
}
