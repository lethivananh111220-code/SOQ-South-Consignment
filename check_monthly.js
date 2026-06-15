const XLSX = require('xlsx');
const workbook = XLSX.readFile('../Doanh số tháng.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
console.log('Headers: ', data[0].join(' | '));
console.log('Row 1: ', data[1] ? data[1].join(' | ') : 'No data');
