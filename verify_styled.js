// Verify styled xlsx files - check data integrity and styling
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_north_single_volume.json'), 'utf-8'));

async function verify() {
  const checks = [
    { rep: "何光傑", customer: "柏興圖書", rawQty: 7541 },
    { rep: "李敏豪", customer: "福源文教", rawQty: 6200 },
    { rep: "康晉瑋", customer: "春秋書坊", rawQty: 5368 },
    { rep: "朱鵬學", customer: "小城書局", rawQty: 4170 },
    { rep: "林智偉", customer: "學園社", rawQty: 5125 },
  ];

  console.log("=== Data Integrity & Styling Verification ===\n");

  for (const { rep, customer, rawQty } of checks) {
    const filePath = path.join(__dirname, 'Output', `${rep}_114上單冊講義.xlsx`);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    
    const ws = wb.getWorksheet(customer);
    if (!ws) { console.log(`✗ ${rep}/${customer} - sheet not found`); continue; }

    // Check styling exists
    const titleCell = ws.getCell(1, 1);
    const hasFill = titleCell.fill && titleCell.fill.fgColor;
    const hasFont = titleCell.font && titleCell.font.bold;
    
    // Sum all numeric data cells to verify total
    let xlsxTotal = 0;
    const sectionLabels = new Set();
    
    ws.eachRow((row, rowNum) => {
      const cellA = row.getCell(1).value;
      const cellB = row.getCell(2).value;
      
      if (['雙向', '735', '試題篇', '新講義'].includes(cellA)) {
        sectionLabels.add(cellA);
      }
      
      // Sum data from version rows (N版/K版/H版/綜合版)
      if (cellB && ['N版', 'K版', 'H版'].includes(String(cellB))) {
        for (let c = 4; c <= 24; c++) {
          const val = row.getCell(c).value;
          if (typeof val === 'number' && val > 0) {
            // Only count subject columns, not summary columns
            // Check if this column is a summary column by looking at header
          }
        }
      }
      
      if (String(cellB).includes('綜合版')) {
        // 新講義 data row
      }
    });
    
    // Simpler: verify by checking 合計 rows
    let totalFromTotals = 0;
    ws.eachRow((row, rowNum) => {
      if (String(row.getCell(2).value) === '合計') {
        // Find the 訂書 column (varies by section)
        // It's the first summary column after subject data
        // For each row, check which section by counting back
        for (let c = 4; c <= 24; c++) {
          const val = row.getCell(c).value;
          if (typeof val === 'number') {
            // The subject cells have individual values, then 訂書 has the sum
            // We need the subject values, not the summary
          }
        }
      }
    });
    
    // Alternative: sum individual version rows' subject cells
    let computedTotal = 0;
    ws.eachRow((row, rowNum) => {
      const cellB = String(row.getCell(2).value || '');
      if (['N版', 'K版', 'H版'].includes(cellB) || cellB.includes('綜合版')) {
        // Find where subject data ends by looking for cells AFTER a sequence of numbers
        // Subject data starts at col 4
        // We need to know how many subject columns this section has
        // Determine section from nearby rows
        const period = String(row.getCell(3).value || '');
        if (period === '114上') {
          // Count all numeric cells from col 4 onwards, but stop before summary columns
          // The summary columns are: 訂書, 退貨, 實際出貨
          // They repeat the total, so we just need the subject cells
          // Simpler: find the total in the row and use that
          // The last non-empty numeric is 實際出貨 which equals total
          
          // Just collect all numeric values, then use the pattern:
          // [subj1, subj2, ..., 訂書, (empty), 實際出貨]
          // 訂書 = sum of subjects, so just take 訂書 value
          const nums = [];
          for (let c = 4; c <= 24; c++) {
            const v = row.getCell(c).value;
            nums.push(typeof v === 'number' ? v : null);
          }
          
          // Find the last non-null number
          let lastIdx = nums.length - 1;
          while (lastIdx >= 0 && nums[lastIdx] === null) lastIdx--;
          
          // The 訂書 value is 3 positions before the end
          // Pattern: ...subjects..., 訂書, null(退貨), 實際出貨
          if (lastIdx >= 2) {
            const actualShipment = nums[lastIdx]; // 實際出貨
            computedTotal += actualShipment || 0;
          }
        }
      }
    });
    
    const dataMatch = computedTotal === rawQty;
    
    console.log(`${dataMatch ? '✓' : '✗'} ${rep}/${customer}`);
    console.log(`  Data: xlsx=${computedTotal}, raw=${rawQty} ${dataMatch ? '✓' : '✗ MISMATCH'}`);
    console.log(`  Styling: fill=${hasFill ? '✓' : '✗'}, font=${hasFont ? '✓' : '✗'}`);
    console.log(`  Sections: ${[...sectionLabels].join(', ')}`);
  }
  
  // Also check overall structure
  console.log("\n=== All Files Structure ===");
  const files = fs.readdirSync(path.join(__dirname, 'Output')).filter(f => f.endsWith('.xlsx'));
  for (const file of files) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, 'Output', file));
    
    let allSheetsOk = true;
    for (const ws of wb.worksheets) {
      // Check title cell has fill
      const cell = ws.getCell(1, 1);
      if (!cell.fill || !cell.fill.fgColor) allSheetsOk = false;
      // Check section headers exist with color
      let hasSections = false;
      ws.eachRow((row) => {
        const v = row.getCell(1).value;
        if (v === '雙向' || v === '735' || v === '試題篇' || v === '新講義') {
          hasSections = true;
        }
      });
      if (!hasSections) allSheetsOk = false;
    }
    
    console.log(`${allSheetsOk ? '✓' : '✗'} ${file} - ${wb.worksheets.length} sheets, styling: ${allSheetsOk ? 'OK' : 'ISSUES'}`);
  }
}

verify().catch(console.error);
