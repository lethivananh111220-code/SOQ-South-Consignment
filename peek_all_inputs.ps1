Add-Type -AssemblyName System.IO.Compression.FileSystem

function Peek-Excel($pattern) {
    Write-Output "--- Pattern: $pattern ---"
    $file = Get-ChildItem -Path "D:\DHF" -Filter $pattern -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($file) {
        Write-Output "File: $($file.FullName)"
        $temp = "d:\DHF\QLKV_WM\web_app\temp_peek_all.xlsx"
        Copy-Item $file.FullName $temp -Force
        
        $zip = [System.IO.Compression.ZipFile]::OpenRead($temp)
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
            $rows = $xml.worksheet.sheetData.row | Select-Object -First 5
            foreach ($r in $rows) {
                $vals = @()
                foreach ($c in $r.c) {
                    $val = ''
                    if ($c.t -eq 's') { 
                        $idx = [int]$c.v
                        if ($idx -lt $strings.Count) {
                            $val = $strings[$idx]
                            if ($val -is [System.Xml.XmlElement]) { $val = $val.InnerText }
                        } else {
                            $val = $c.v
                        }
                    } else {
                        $val = $c.v
                    }
                    $vals += $val
                }
                Write-Output "Row $($r.r): $($vals -join ' | ')"
            }
        }
        $zip.Dispose()
        Remove-Item $temp -Force
    } else {
        Write-Output "No file found for pattern: $pattern"
    }
    Write-Output ""
}

Peek-Excel "merchandiser_report*"
Peek-Excel "Monthly revenue*"
Peek-Excel "Sell-Report-Download*"
