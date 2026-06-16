Add-Type -AssemblyName System.IO.Compression.FileSystem
$parentDir = "D:\DHF\CONSIGNMENT\Don giao 0104"
$file = Get-ChildItem -Path $parentDir -Filter "*2303-0504*" | Select-Object -First 1

if ($file) {
    Write-Output "Processing: $($file.FullName)"
    $tempPath = "d:\DHF\QLKV_WM\web_app\temp_check.xlsx"
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
        
        # Find column for "1/4"
        $targetColIdx = -1
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $h = $headers[$i]
            if ($h -match "^1$|^1/4$|^01/04$|^1.4$|^1/04$") {
                $targetColIdx = $i
                Write-Output "Found target column: '$h' at index $i"
            }
        }
        
        if ($targetColIdx -ne -1) {
            # Find store name/SAP column
            $sapColIdx = 0
            $storeColIdx = 1
            for ($i = 0; $i -lt $headers.Count; $i++) {
                if ($headers[$i] -match "SAP|Store Key|Ma Kho|Mach") { $sapColIdx = $i }
                if ($headers[$i] -match "Store Name|Ten Cua Hang|Tncahng") { $storeColIdx = $i }
            }

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
                
                $deliverySignal = $rowValues[$targetColIdx]
                if ($deliverySignal -and $deliverySignal.Trim() -ne "" -and $deliverySignal -ne "0" -and $deliverySignal -notmatch "nghỉ|off") {
                    $sap = $rowValues[$sapColIdx]
                    $store = $rowValues[$storeColIdx]
                    Write-Output "[$sap] $store : $deliverySignal"
                }
            }
        } else {
            Write-Output "Headers found: $($headers -join ' | ')"
        }
    }
    $zip.Dispose()
    Remove-Item $tempPath -Force
} else {
    Write-Output "File not found"
}
