Add-Type -AssemblyName System.IO.Compression.FileSystem
$f = "D:\DHF\CONSIGNMENT\Don giao 0104\Lịch 2303-0504.xlsx" # Exact path identified earlier
if (-not (Test-Path $f)) {
    # Fallback to searching if the literal path fails due to special characters
    $f = (Get-ChildItem -Path "D:\DHF\CONSIGNMENT\Don giao 0104" -Filter "*2303-0504*" | Select-Object -First 1).FullName
}

Write-Output "Checking file: $f"

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
    $rows = $xml.worksheet.sheetData.row
    
    # Get Headers (assuming row 1)
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
    
    # Find column for "1/4" or "1.04" or "1"
    $targetColIdx = -1
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $h = $headers[$i]
        if ($h -match "^1$|^1/4$|^01/04$|^1.4$|^1/04$") {
            $targetColIdx = $i
            Write-Output "Found target column: '$h' at index $i"
        }
    }
    
    if ($targetColIdx -eq -1) {
        Write-Output "Could not find a column for April 1st (1/4)."
    } else {
        Write-Output "Stores scheduled for delivery on 1.4:"
        # Find store name/SAP column
        $sapColIdx = 0
        $storeColIdx = 1
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i] -match "SAP|Store Key|Ma Kho|Mach") { $sapColIdx = $i }
            if ($headers[$i] -match "Store Name|Ten Cua Hang|Tncahng") { $storeColIdx = $i }
        }

        foreach ($r in ($rows | Select-Object -Skip 1)) {
            $cols = @()
            # Map cells by their reference (A1, B1...) because columns might be missing
            # But the row object has 'c' elements. We need to be careful with missing columns.
            $rowValues = @{}
            foreach ($c in $r.c) {
                # Column index from reference (e.g., A1 -> 0, B1 -> 1, AA1 -> 26)
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
    }
}
$zip.Dispose()
