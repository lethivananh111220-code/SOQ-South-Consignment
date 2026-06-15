try {
    var fs = new ActiveXObject('Scripting.FileSystemObject');
    var code = fs.OpenTextFile('d:/DHF/SOQ - H? N?I/website/app.js', 1).ReadAll();
    eval('function test() { ' + code + ' }');
    WScript.Echo('Syntax OK');
} catch (e) {
    WScript.Echo('Syntax Error: ' + e.description + ' at line ' + e.line);
}
