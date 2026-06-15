$code = [System.IO.File]::ReadAllText('d:/DHF/SOQ - HÀ NỘI/website/app.js', [System.Text.Encoding]::UTF8)
try {
    $sc = New-Object -ComObject MSScriptControl.ScriptControl
    $sc.Language = 'JScript'
    $sc.AddCode($code)
    Write-Host 'Syntax OK'
} catch {
    Write-Host 'Syntax Error:' $_
}
