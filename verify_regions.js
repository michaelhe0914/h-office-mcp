import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const filePath = path.join(__dirname, 'Output', '各區單冊講義產品總數量表.xlsx');
const jsonPath = path.join(__dirname, 'query_results_all_regions.json');

async function verify() {
  console.log("=== STARTING REGIONAL XLSX VERIFICATION ===\n");
  
  if (!fs.existsSync(filePath)) {
    console.error(`✗ Error: Excel file not found at: ${filePath}`);
    process.exit(1);
  }
  
  if (!fs.existsSync(jsonPath)) {
    console.error(`✗ Error: JSON data file not found at: ${jsonPath}`);
    process.exit(1);
  }
  
  // Load JSON raw total
  const rawRecords = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
  const rawTotalQty = rawRecords.reduce((sum, r) => sum + r.qty, 0);
  console.log(`Raw JSON total qty (all regions): ${rawTotalQty}`);
  
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  
  // 1. Verify sheet counts and names
  console.log("\n1. Verifying sheet names and order:");
  const expectedSheets = ['總表', '北區', '中區', '南區'];
  if (wb.worksheets.length !== 4) {
    console.log(`  ✗ Expected 4 sheets, found ${wb.worksheets.length}`);
    process.exit(1);
  }
  for (let i = 0; i < 4; i++) {
    const ws = wb.worksheets[i];
    const pass = ws.name === expectedSheets[i];
    console.log(`  ${pass ? '✓' : '✗'} Sheet index ${i}: '${ws.name}' (expected '${expectedSheets[i]}')`);
    if (!pass) process.exit(1);
  }
  
  // Helper to extract numeric data from standard styled sheet rows
  function extractSheetDataTotals(ws) {
    let grandSum = 0;
    
    // Find where the summary table starts
    let sumTitleRow = -1;
    ws.eachRow((row, rowNum) => {
      if (row.getCell(1).value === '各品項及年級加總表') {
        sumTitleRow = rowNum;
      }
    });
    
    if (sumTitleRow === -1) {
      console.log(`  ✗ ${ws.name}: '各品項及年級加總表' not found!`);
      return null;
    }
    
    // Check that we can read the product totals by summing version rows' subject cells
    // Version rows are K版, N版, H版 or 綜合版 合計 (or contains 綜合版)
    // We sum all subject numeric cells.
    // Let's identify the rows:
    ws.eachRow((row, rowNum) => {
      const cellB = String(row.getCell(2).value || '');
      if (['N版', 'K版', 'H版'].includes(cellB) || cellB.includes('綜合版')) {
        const period = String(row.getCell(3).value || '');
        if (period === '114上') {
          // Find the last numeric cell in the row which represents the row total (實際出貨)
          const nums = [];
          for (let c = 4; c <= 24; c++) {
            const v = row.getCell(c).value;
            nums.push(typeof v === 'number' ? v : null);
          }
          let lastIdx = nums.length - 1;
          while (lastIdx >= 0 && nums[lastIdx] === null) lastIdx--;
          if (lastIdx >= 2) {
            grandSum += nums[lastIdx];
          }
        }
      }
    });
    
    return { grandSum, sumTitleRow };
  }
  
  // 2. Extract totals from each sheet
  console.log("\n2. Verifying sheet data totals:");
  const sheetTotals = {};
  for (const zone of expectedSheets) {
    const ws = wb.getWorksheet(zone);
    const res = extractSheetDataTotals(ws);
    if (!res) {
      console.log(`  ✗ Failed to extract data for ${zone}`);
      process.exit(1);
    }
    sheetTotals[zone] = res.grandSum;
    console.log(`  ✓ ${zone} sheet total data: ${res.grandSum}`);
  }
  
  // Compare sum of regions with 總表
  const sumOfRegions = sheetTotals['北區'] + sheetTotals['中區'] + sheetTotals['南區'];
  const totalMatch = sumOfRegions === sheetTotals['總表'];
  console.log(`  ${totalMatch ? '✓' : '✗'} Sum of (北區 + 中區 + 南區) = ${sumOfRegions}, 總表 = ${sheetTotals['總表']}`);
  if (!totalMatch) process.exit(1);
  
  // Compare 總表 with JSON raw total
  const jsonMatch = sheetTotals['總表'] === rawTotalQty;
  console.log(`  ${jsonMatch ? '✓' : '✗'} Excel 總表 = ${sheetTotals['總表']}, JSON Raw = ${rawTotalQty}`);
  if (!jsonMatch) process.exit(1);
  
  // 3. Detailed Cell-by-Cell Summation Check:
  // For each category, version, subject, and grade, check that (北區 + 中區 + 南區) === 總表
  console.log("\n3. Cross-checking cell-by-cell regional sum values:");
  const wsTotal = wb.getWorksheet('總表');
  const wsNorth = wb.getWorksheet('北區');
  const wsCentral = wb.getWorksheet('中區');
  const wsSouth = wb.getWorksheet('南區');
  
  let cellMismatch = false;
  // Let's iterate all rows and columns to find numeric data cells (subject columns in data rows)
  for (let r = 1; r <= wsTotal.rowCount; r++) {
    const rowT = wsTotal.getRow(r);
    const rowN = wsNorth.getRow(r);
    const rowC = wsCentral.getRow(r);
    const rowS = wsSouth.getRow(r);
    
    const cellB_Val = String(rowT.getCell(2).value || '');
    if (['N版', 'K版', 'H版'].includes(cellB_Val) || cellB_Val.includes('綜合版')) {
      const period = String(rowT.getCell(3).value || '');
      if (period === '114上') {
        // This is a data row. Check subject columns.
        // Subject columns vary depending on section.
        // Let's find how many subject columns are in this row by looking at numeric cells before summary columns.
        // Let's just check columns 4 to 21 (subject columns)
        for (let c = 4; c <= 21; c++) {
          const valT = rowT.getCell(c).value || 0;
          const valN = rowN.getCell(c).value || 0;
          const valC = rowC.getCell(c).value || 0;
          const valS = rowS.getCell(c).value || 0;
          
          if (typeof valT === 'number' || typeof valN === 'number' || typeof valC === 'number' || typeof valS === 'number') {
            const sum = (typeof valN === 'number' ? valN : 0) + 
                        (typeof valC === 'number' ? valC : 0) + 
                        (typeof valS === 'number' ? valS : 0);
            const expected = typeof valT === 'number' ? valT : 0;
            if (sum !== expected) {
              console.log(`  ✗ Cell Mismatch at Row ${r} Col ${c}: 北(${valN}) + 中(${valC}) + 南(${valS}) = ${sum}, 總表 = ${expected}`);
              cellMismatch = true;
            }
          }
        }
      }
    }
  }
  
  if (cellMismatch) {
    console.log("  ✗ Cell-by-cell verification failed!");
    process.exit(1);
  } else {
    console.log("  ✓ All numeric data cells sum up perfectly across regions!");
  }
  
  // 4. Verification of Formulas and Visual Boundaries
  console.log("\n4. Verification of summary table formulas & visual boundaries:");
  for (const zone of expectedSheets) {
    const ws = wb.getWorksheet(zone);
    const { sumTitleRow } = extractSheetDataTotals(ws);
    
    // Check summary table formulas exist
    const categories = ['雙向', '735', '試題篇', '新講義', '合計'];
    let formulasOk = true;
    for (let i = 0; i < categories.length; i++) {
      const rNum = sumTitleRow + 2 + i;
      const row = ws.getRow(rNum);
      for (let c = 2; c <= 5; c++) {
        const cell = row.getCell(c);
        if (!cell.value || typeof cell.value !== 'object' || !cell.value.formula) {
          formulasOk = false;
        }
      }
    }
    console.log(`  ✓ ${zone}: Summary table formulas present: ${formulasOk ? 'YES' : 'NO'}`);
    if (!formulasOk) process.exit(1);
    
    // Check for border checks in cols 6-24 of the summary table rows
    let borderError = false;
    for (let r = sumTitleRow; r <= sumTitleRow + 6; r++) {
      for (let c = 6; c <= 24; c++) {
        const cell = ws.getCell(r, c);
        if (cell.border && (cell.border.top || cell.border.left || cell.border.bottom || cell.border.right)) {
          borderError = true;
        }
      }
    }
    console.log(`  ✓ ${zone}: No hanging borders to the right: ${!borderError ? 'YES' : 'NO'}`);
    if (borderError) process.exit(1);
  }
  
  console.log("\n=== REGIONAL XLSX VERIFICATION SUCCESSFULLY COMPLETED! ALL CHECKS PASSED ✓ ===");
}

verify().catch(e => {
  console.error(e);
  process.exit(1);
});
