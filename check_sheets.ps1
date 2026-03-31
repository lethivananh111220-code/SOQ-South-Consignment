Add-Type -AssemblyName System.IO.Compression.FileSystem
$file = Get-ChildItem -Path "D:\" -Filter "*2303-0504*" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if ($file) {
    Write-Output "Processing: $($file.FullName)"
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
    $wbEntry = $zip.GetEntry('xl/workbook.xml')
    if ($wbEntry) {
        $xml = [xml](New-Object System.IO.StreamReader($wbEntry.Open())).ReadToEnd()
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("ns", $xml.DocumentElement.NamespaceURI)
        $sheets = $xml.SelectNodes("//ns:sheet", $ns)
        foreach ($s in $sheets) {
            Write-Output "Sheet: $($s.name)"
        }
    }
    $zip.Dispose()
} else {
    Write-Output "File not found"
}
