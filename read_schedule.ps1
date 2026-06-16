Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-XlsxHeaders {
    param($resolvedPath)
    if (-not (Test-Path $resolvedPath)) { return }
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
        $rows = $xml.worksheet.sheetData.row | Select-Object -First 3
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

$source = Get-Item "d:\DHF\QLKV_WM\ĐƠN ĐẶT HÀNG\Đơn đặt ngày 1903\Lịch 2003-2203.xlsx" -ErrorAction SilentlyContinue
if (-not $source) {
    # Try alternate find just in case
    $source = Get-ChildItem "d:\DHF\QLKV_WM" -Recurse | Where-Object { $_.Name -match 'Lịch' } | Select-Object -First 1
}

if ($source) {
    Write-Output "File: $($source.FullName)"
    $tempTarget = 'd:\DHF\QLKV_WM\web_app\temp_sched2.xlsx'
    Copy-Item $source.FullName $tempTarget -Force
    Get-XlsxHeaders $tempTarget
    Remove-Item $tempTarget -Force
} else {
    Write-Output "STILL NOT FOUND"
}
