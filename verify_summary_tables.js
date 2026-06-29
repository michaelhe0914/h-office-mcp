import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

// Target sales reps
const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉瑋", "朱鵬學", "蔡榮訓", "曹原菘", "林廣強"];
const outputDir = 'd:/Antigravity/出書統計表/Output';

async function verifyAll() {
  console.log("=== VERIFYING SUMMARY TABLES IN EXCEL FILES ===\n");
  
  let allOk = true;

  for (const rep of targetSales) {
    const fileName = `${rep}_114上單冊講義.xlsx`;
    const filePath = path.join(outputDir, fileName);
    
    if (!fs.existsSync(filePath)) {
      console.log(`✗ File not found for ${rep}: ${fileName}`);
      allOk = false;
      continue;
    }
    
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    console.log(`Checking ${fileName} (${wb.worksheets.length} sheets)...`);
    
    for (const ws of wb.worksheets) {
      // Find the row containing "各品項及年級加總表"
      let sumTitleRow = -1;
      ws.eachRow((row, rowNum) => {
        if (row.getCell(1).value === '各品項及年級加總表') {
          sumTitleRow = rowNum;
        }
      });
      
      if (sumTitleRow === -1) {
        console.log(`  ✗ ${ws.name}: '各品項及年級加總表' not found!`);
        allOk = false;
        continue;
      }
      
      // The headers should be: 品項, 一年級, 三年級, 五年級, 總本數
      const headerRow = ws.getRow(sumTitleRow + 1);
      const expectedHeaders = ['品項', '一年級', '三年級', '五年級', '總本數'];
      for (let c = 1; c <= 5; c++) {
        const val = headerRow.getCell(c).value;
        if (val !== expectedHeaders[c - 1]) {
          console.log(`  ✗ ${ws.name}: Header col ${c} is '${val}', expected '${expectedHeaders[c-1]}'`);
          allOk = false;
        }
      }
      
      // Check formulas for each row (雙向, 735, 試題篇, 新講義, 合計)
      const categories = ['雙向', '735', '試題篇', '新講義', '合計'];
      for (let i = 0; i < categories.length; i++) {
        const rNum = sumTitleRow + 2 + i;
        const row = ws.getRow(rNum);
        
        // Col 1 must be category name
        if (row.getCell(1).value !== categories[i]) {
          console.log(`  ✗ ${ws.name}: Row ${rNum} col 1 is '${row.getCell(1).value}', expected '${categories[i]}'`);
          allOk = false;
        }
        
        // Col 2-5 must have formulas
        for (let c = 2; c <= 5; c++) {
          const cell = row.getCell(c);
          if (!cell.value || typeof cell.value !== 'object' || !cell.value.formula) {
            console.log(`  ✗ ${ws.name}: Cell at row ${rNum} col ${c} does not contain a formula! Value:`, cell.value);
            allOk = false;
          }
        }
      }
      
      // Let's print out the formulas of the first sheet of each rep as a sample
      if (ws === wb.worksheets[0]) {
        console.log(`  Sample formulas for sheet "${ws.name}":`);
        for (let i = 0; i < categories.length; i++) {
          const rNum = sumTitleRow + 2 + i;
          const row = ws.getRow(rNum);
          const f2 = row.getCell(2).value.formula;
          const f3 = row.getCell(3).value.formula;
          const f4 = row.getCell(4).value.formula;
          const f5 = row.getCell(5).value.formula;
          console.log(`    ${categories[i].padEnd(6)} | 一年級: =${f2.padEnd(20)} | 三年級: =${f3.padEnd(20)} | 五年級: =${f4.padEnd(20)} | 總本數: =${f5}`);
        }
        
        // Check that cells in cols 6-24 for the summary table rows DO NOT have borders
        for (let r = sumTitleRow; r <= sumTitleRow + 6; r++) {
          for (let c = 6; c <= 24; c++) {
            const cell = ws.getCell(r, c);
            if (cell.border && (cell.border.top || cell.border.left || cell.border.bottom || cell.border.right)) {
              console.log(`  ✗ ${ws.name}: Hanging border found at row ${r} col ${c}!`);
              allOk = false;
            }
          }
        }
        console.log(`  ✓ Sample border check passed (no hanging borders on the right).`);
      }
    }
  }

  // Verify regional grand summary files
  const regionalFiles = [
    { name: '北區單冊講義總數量合計總表.xlsx', title: '北區總數量合計總表' },
    { name: '中區單冊講義總數量合計總表.xlsx', title: '中區總數量合計總表' }
  ];

  for (const reg of regionalFiles) {
    console.log(`\nChecking ${reg.name}...`);
    const grandPath = path.join(outputDir, reg.name);
    if (!fs.existsSync(grandPath)) {
      console.log(`✗ Grand summary file not found: ${grandPath}`);
      allOk = false;
      continue;
    }
    
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(grandPath);
    if (wb.worksheets.length !== 1) {
      console.log(`  ✗ Expected 1 worksheet, found ${wb.worksheets.length}`);
      allOk = false;
      continue;
    }
    const ws = wb.worksheets[0];
    if (ws.name !== '總數量合計總表') {
      console.log(`  ✗ Expected sheet name '總數量合計總表', found '${ws.name}'`);
      allOk = false;
      continue;
    }
    
    // Perform standard summary table checks on grand summary
    let sumTitleRow = -1;
    ws.eachRow((row, rowNum) => {
      if (row.getCell(1).value === '各品項及年級加總表') {
        sumTitleRow = rowNum;
      }
    });
    
    if (sumTitleRow === -1) {
      console.log(`  ✗ ${reg.name}: '各品項及年級加總表' not found!`);
      allOk = false;
    } else {
      const headerRow = ws.getRow(sumTitleRow + 1);
      const expectedHeaders = ['品項', '一年級', '三年級', '五年級', '總本數'];
      for (let c = 1; c <= 5; c++) {
        const val = headerRow.getCell(c).value;
        if (val !== expectedHeaders[c - 1]) {
          console.log(`  ✗ ${reg.name}: Header col ${c} is '${val}', expected '${expectedHeaders[c-1]}'`);
          allOk = false;
        }
      }
      
      const categories = ['雙向', '735', '試題篇', '新講義', '合計'];
      for (let i = 0; i < categories.length; i++) {
        const rNum = sumTitleRow + 2 + i;
        const row = ws.getRow(rNum);
        if (row.getCell(1).value !== categories[i]) {
          console.log(`  ✗ ${reg.name}: Row ${rNum} col 1 is '${row.getCell(1).value}', expected '${categories[i]}'`);
          allOk = false;
        }
        for (let c = 2; c <= 5; c++) {
          const cell = row.getCell(c);
          if (!cell.value || typeof cell.value !== 'object' || !cell.value.formula) {
            console.log(`  ✗ ${reg.name}: Cell at row ${rNum} col ${c} does not contain formula!`);
            allOk = false;
          }
        }
      }
      console.log(`  ✓ ${reg.name} structure & formulas OK.`);
    }
  }
  
  if (allOk) {
    console.log("\n=== ALL FILES SUCCESSFULLY VERIFIED! ALL TESTS PASSED ✓ ===");
  } else {
    console.log("\n=== SOME VERIFICATION CHECKS FAILED ✗ ===");
  }
}

verifyAll().catch(console.error);
