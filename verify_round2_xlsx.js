import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function verify() {
  const dir = path.join(__dirname, '第2波複習1');
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.xlsx'));

  console.log(`Found ${files.length} excel files in 第2波複習1:`);
  for (const f of files) {
    const filePath = path.join(dir, f);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);

    console.log(`\nFile: ${f}`);
    console.log(`  Sheets Count: ${wb.worksheets.length}`);

    // Verify first sheet
    const ws1 = wb.worksheets[0];
    console.log(`  Sheet 1 Name: ${ws1.name}`);
    console.log(`  Sheet 1 Rows: ${ws1.rowCount}, Columns: ${ws1.columnCount}`);

    // If it's overall overview sheet, print row contents
    if (ws1.name === '業務與客戶統計') {
      console.log(`  Overview Sheet Data:`);
      ws1.eachRow((row, rowNumber) => {
        if (rowNumber >= 2) {
          const vals = row.values.slice(1, 8).map(v => typeof v === 'object' ? JSON.stringify(v) : v);
          console.log(`    Row ${rowNumber}: ${vals.join(' | ')}`);
        }
      });
    }
  }
}

verify().catch(console.error);
