const XLSX = require('d:/DHF/QLKV_WM/web_app/node_modules/xlsx');

function normalizeKey(key) {
    if (!key) return '';
    return key.toString().toLowerCase().replace(/[^a-z0-9\/\-]/g, '');
}

function extractJsonDataCleanly(worksheet) {
    const rawArr = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    let headerIdx = 0;
    for (let i = 0; i < Math.min(20, rawArr.length); i++) {
        let r = rawArr[i];
        let validCols = r.filter(c => c !== null && c !== undefined && c !== "");
        if (validCols.length >= 2 && r.some(c => typeof c === 'string' && (c.toUpperCase().includes('SAP') || c.toUpperCase().includes('STORE') || c.toUpperCase().includes('NICKNAME') || c.toUpperCase().includes('TÊN')))) {
            headerIdx = i;
            break;
        }
    }

    const headersRaw = rawArr[headerIdx] || [];
    let headers = headersRaw.map(h => normalizeKey(h));
    
    let numericHeadersCount = headersRaw.filter(h => typeof h === 'number' && h > 40000).length;
    if (numericHeadersCount > (headersRaw.length / 2)) {
        headers[0] = 'type';
        headers[1] = 'sap';
        headers[4] = 'storename';
    }

    const json = [];
    for (let i = headerIdx + 1; i < rawArr.length; i++) {
        let row = rawArr[i];
        if (!row || row.length === 0) continue;
        let obj = {};
        let hasData = false;
        for (let j = 0; j < row.length; j++) {
            let val = row[j];
            if (val !== "" && val !== null && val !== undefined) {
                obj[headers[j] || ('col' + j)] = val;
                obj[headersRaw[j]] = val;
                hasData = true;
            }
        }
        if (hasData) json.push(obj);
    }
    return json;
}

const file = 'D:/DHF/CONSIGNMENT/Don giao 0104/Lịch 2303-0504.xlsx'.replace(/Lịch/, 'L?ch'); // Handle the ?
// Wait, I should find it properly
const fs = require('fs');
const path = require('path');
const dir = 'D:/DHF/CONSIGNMENT/Don giao 0104';
const files = fs.readdirSync(dir);
const targetFile = files.find(f => f.includes('2303-0504'));
const fullPath = path.join(dir, targetFile);

console.log('Reading:', fullPath);
const workbook = XLSX.readFile(fullPath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = extractJsonDataCleanly(sheet);

console.log('Total rows:', data.length);
const s6295 = data.find(r => r.sap == '6295' || r.nickname == '6295' || r['46105'] == '6295');
console.log('Store 6295:', JSON.stringify(s6295, null, 2));

const s5591 = data.find(r => r.sap == '5591' || r.nickname == '5591' || r['46105'] == '5591');
console.log('Store 5591:', JSON.stringify(s5591, null, 2));
