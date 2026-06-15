try {
    $sc = New-Object -ComObject MSScriptControl.ScriptControl
    $sc.Language = 'JScript'
    $sc.AddCode('let x = 1;')
    Write-Host 'Syntax OK'
} catch {
    Write-Host 'Syntax Error:' $_
}
