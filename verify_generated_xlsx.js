import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function verify() {
  const masterPath = path.join(__dirname, '半全冊', '2025.7.01-9.25_北區半全冊卷出書統計表.xlsx');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(masterPath);

  console.log(`=== Verification for ${masterPath} ===`);
  console.log(`Total Worksheets: ${wb.worksheets.length}`);
  wb.worksheets.forEach((ws, i) => console.log(`  Sheet ${i + 1}: ${ws.name} (Rows: ${ws.rowCount})`));

  // Verify Sheet 1: 北區業務與客戶統計
  const ws1 = wb.getWorksheet('北區業務與客戶統計');
  console.log('\n--- Sheet 1: 北區業務與客戶統計 Content Check ---');
  ws1.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const vals = row.values.slice(1);
    if (rowNumber <= 15) {
      console.log(`Row ${rowNumber}:`, vals.map(v => typeof v === 'object' ? (v.result !== undefined ? `FormulaResult:${v.result}` : v.formula) : v));
    }
  });

  // Verify Sheet 2: 半全冊卷全區產品總彙
  const ws2 = wb.getWorksheet('半全冊卷全區產品總彙');
  console.log('\n--- Sheet 2: 半全冊卷全區產品總彙 Check ---');
  console.log(`Product count rows: ${ws2.rowCount}`);

  // Verify Sheet 3: 北區全區出貨總明細
  const ws3 = wb.getWorksheet('北區全區出貨總明細');
  console.log('\n--- Sheet 3: 北區全區出貨總明細 Check ---');
  console.log(`Detail count rows: ${ws3.rowCount}`);

  console.log('\n✓ Excel file structural verification complete!');
}

verify().catch(console.error);
