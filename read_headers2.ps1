Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-XlsxRow {
    param($pathPattern)
    $resolvedPaths = @(Resolve-Path $pathPattern -ErrorAction SilentlyContinue)
    if ($resolvedPaths.Count -eq 0) {
        Write-Output "File not found: $pathPattern"
        return
    }
    $resolvedPath = $resolvedPaths[0].Path
    
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

Write-Output '--- WEEKLY ---'
Get-XlsxRow 'd:\DHF\QLKV_WM\*\*\W11*.xlsx'

Write-Output '--- INPUT ---'
Get-XlsxRow 'd:\DHF\QLKV_WM\*\*\Sell-Report*.xlsx'
