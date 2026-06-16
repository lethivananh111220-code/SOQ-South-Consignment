Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-XlsxHeaders {
    param($resolvedPath)
    if (-not (Test-Path $resolvedPath)) {
        Write-Output "File not found: $resolvedPath"
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
        $rows = $xml.worksheet.sheetData.row | Select-Object -First 10
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

$source = Get-ChildItem 'd:\DHF\QLKV_WM\*\*\Sell-Report*.xlsx' | Select-Object -First 1
$tempTarget = 'd:\DHF\QLKV_WM\web_app\temp_sell.xlsx'
Copy-Item $source.FullName $tempTarget -Force

Write-Output '--- INPUT ODA ---'
Get-XlsxHeaders $tempTarget
Remove-Item $tempTarget -Force
