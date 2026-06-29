import XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Verify one of the generated files
const filePath = path.join(__dirname, 'Output', '何光傑_114上單冊講義.xlsx');
const workbook = XLSX.readFile(filePath);

console.log("=== Sheet Names ===");
console.log(workbook.SheetNames);

// Check the first sheet (柏興圖書 - has highest qty)
const sheetName = '柏興圖書';
const sheet = workbook.Sheets[sheetName];

console.log(`\n=== Sheet: "${sheetName}" ===`);
const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
for (let i = 0; i < jsonData.length; i++) {
  const row = jsonData[i];
  if (row.some(v => v !== '' && v !== null && v !== undefined)) {
    console.log(`Row ${i}: ${JSON.stringify(row)}`);
  }
}

// Also verify reference format
console.log("\n\n=== Reference Format (中區) ===");
const refPath = path.join(__dirname, 'Input', '114上各單冊講義比較.xlsx');
const refWb = XLSX.readFile(refPath);
const refSheet = refWb.Sheets['中區'];
const refData = XLSX.utils.sheet_to_json(refSheet, { header: 1, defval: '' });
for (let i = 0; i < refData.length; i++) {
  const row = refData[i];
  if (row.some(v => v !== '' && v !== null && v !== undefined)) {
    console.log(`Row ${i}: ${JSON.stringify(row)}`);
  }
}
