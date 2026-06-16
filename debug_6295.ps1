Add-Type -AssemblyName System.IO.Compression.FileSystem
$file = Get-ChildItem -Path "D:\" -Filter "*2303-0504*" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if ($file) {
    Write-Output "Processing: $($file.FullName)"
    $tempPath = "d:\DHF\QLKV_WM\web_app\temp_debug_6295.xlsx"
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
        $rows = $xml.worksheet.sheetData.row
        
        # Get Headers
        $headerRow = $rows | Select-Object -First 1
        $headers = @()
        foreach ($c in $headerRow.c) {
            $val = ""
            if ($c.t -eq 's') { 
                $val = $strings[[int]$c.v]
                if ($val -is [System.Xml.XmlElement]) { $val = $val.InnerText }
            } else {
                $val = $c.v
            }
            $headers += $val
        }
        
        Write-Output "Headers: $($headers -join ' | ')"
        
        # Find row for 6295
        $found = $false
        foreach ($r in ($rows | Select-Object -Skip 1)) {
            $rowValues = @{}
            foreach ($c in $r.c) {
                $ref = $c.r -replace '\d+'
                $idx = 0
                for ($j = 0; $j -lt $ref.Length; $j++) {
                    $idx = $idx * 26 + ([int][char]$ref[$j] - [int][char]'A' + 1)
                }
                $idx -= 1
                
                $val = ""
                if ($c.t -eq 's') { 
                    $val = $strings[[int]$c.v]
                    if ($val -is [System.Xml.XmlElement]) { $val = $val.InnerText }
                } else {
                    $val = $c.v
                }
                $rowValues[$idx] = $val
            }
            
            if ($rowValues.Values -contains "6295") {
                $found = $true
                Write-Output "DEBUG (Store 6295 row):"
                foreach ($k in ($rowValues.Keys | Sort-Object)) {
                    Write-Output "Col $k ($($headers[$k])): $($rowValues[$k])"
                }
            }
        }
        if (-not $found) { Write-Output "Store 6295 not found in schedule file." }
    }
    $zip.Dispose()
    Remove-Item $tempPath -Force
} else {
    Write-Output "File not found"
}
