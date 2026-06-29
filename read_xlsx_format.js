import XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const filePath = path.join(__dirname, 'Input', '114上各單冊講義比較.xlsx');

const workbook = XLSX.readFile(filePath);

for (const sheetName of workbook.SheetNames) {
  console.log(`\n=== Sheet: "${sheetName}" ===`);
  const sheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  console.log(`Range: ${sheet['!ref']}`);

  // Print all data as JSON rows
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  for (let i = 0; i < jsonData.length; i++) {
    // Only print rows that have at least some non-empty data
    const row = jsonData[i];
    if (row.some(v => v !== '' && v !== null && v !== undefined)) {
      console.log(`Row ${i}: ${JSON.stringify(row)}`);
    }
  }

  // Print merges
  if (sheet['!merges']) {
    console.log(`\nMerged cells:`);
    for (const merge of sheet['!merges']) {
      const s = XLSX.utils.encode_cell(merge.s);
      const e = XLSX.utils.encode_cell(merge.e);
      console.log(`  ${s}:${e}`);
    }
  }
}
