Add-Type -AssemblyName System.IO.Compression.FileSystem
$file = Get-ChildItem -Path "D:\" -Filter "*2303-0504*" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if ($file) {
    Write-Output "Processing: $($file.FullName)"
    $tempPath = "d:\DHF\QLKV_WM\web_app\temp_peek_v2.xlsx"
    Copy-Item $file.FullName $tempPath -Force
    
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
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
        $rows = $xml.worksheet.sheetData.row | Select-Object -First 10
        foreach ($r in $rows) {
            $vals = @()
            foreach ($c in $r.c) {
                $val = ''
                if ($c.t -eq 's') { 
                    $val = $strings[[int]$c.v]
                    if ($val -is [System.Xml.XmlElement]) { $val = $val.InnerText }
                } else {
                    $val = $c.v
                }
                $vals += $val
            }
            Write-Output "Row $($r.r): $($vals -join ' | ')"
        }
    }
    $zip.Dispose()
    Remove-Item $tempPath -Force
} else {
    Write-Output "File not found"
}
