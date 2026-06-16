Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-XlsxRow {
    param($path)
    $resolvedPath = (Resolve-Path $path).Path
    if (-not (Test-Path $resolvedPath)) {
        Write-Output "File not found: $path"
        return
    }
    
    $zip = [System.IO.Compression.ZipFile]::OpenRead($resolvedPath)
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
        $rows = $xml.worksheet.sheetData.row | Select-Object -First 2
        foreach ($r in $rows) {
            $rowValues = @()
            foreach ($c in $r.c) {
                if ($c.t -eq 's') { $rowValues += $strings[[int]$c.v] }
                else { $rowValues += $c.v }
            }
            Write-Output ($rowValues -join '|')
        }
    }
    $zip.Dispose()
}

Write-Output '--- MAPPING ---'
Get-XlsxRow 'd:\DHF\QLKV_WM\*\Mapping*.xlsx'

Write-Output '--- MONTHLY ---'
Get-XlsxRow 'd:\DHF\QLKV_WM\*\*\Monthly*.xlsx'

Write-Output '--- INVENTORY ---'
Get-XlsxRow 'd:\DHF\QLKV_WM\*\*\merchandiser_report*.xlsx'
