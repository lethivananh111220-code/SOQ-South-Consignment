Add-Type -AssemblyName System.IO.Compression.FileSystem

$dir = "d:\DHF\QLKV_WM\ĐƠN ĐẶT HÀNG\Đơn đặt ngày 1903"
$files = Get-ChildItem -Path $dir -Filter "*.xlsx"

foreach ($f in $files) {
    Write-Output "--- $($f.Name) ---"
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($f.FullName)
        
        $sharedStrEntry = $zip.GetEntry("xl/sharedStrings.xml")
        $strings = @()
        if ($sharedStrEntry) {
            $stream = $sharedStrEntry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $content = $reader.ReadToEnd()
            $xml = [xml]$content
            $strings = $xml.sst.si | ForEach-Object {
                if ($_.t) { $_.t } else { $_.InnerText }
            }
            $reader.Close()
        }
        
        $sheetEntry = $zip.GetEntry("xl/worksheets/sheet1.xml")
        if (-not $sheetEntry) {
            $sheetEntry = $zip.GetEntry("xl/worksheets/Sheet1.xml")
        }
        if ($sheetEntry) {
            $stream = $sheetEntry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $content = $reader.ReadToEnd()
            $xml = [xml]$content
            $row1 = $xml.worksheet.sheetData.row | Select-Object -First 1
            $headers = @()
            if ($row1 -and $row1.c) {
                foreach ($c in $row1.c) {
                    if ($c.t -eq "s") {
                        $idx = [int]$c.v
                        if ($idx -lt $strings.Count) {
                            $val = $strings[$idx]
                            # Sometimes innerText holds the full string
                            if ($val -is [System.Xml.XmlElement]) { $val = $val.InnerText }
                            $headers += $val
                        } else {
                            $headers += $c.v
                        }
                    } else {
                        $headers += $c.v
                    }
                }
            }
            Write-Output ($headers -join ", ")
            $reader.Close()
        } else {
            Write-Output "Sheet1.xml not found"
        }
        
        $zip.Dispose()
    } catch {
        Write-Output "Error parsing $($f.Name): $($_.Exception.Message)"
    }
    Write-Output ""
}
