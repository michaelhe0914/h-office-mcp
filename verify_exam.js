// Verify exam paper xlsx data against JSON source
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_north_exam.json'), 'utf-8'));
const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉瑋", "朱鵬學"];

// Replicate the parsing logic from generate_exam_xlsx.js
function parseExamProduct(product, productClass) {
  const volMatch = product.match(/[（(]([１-６1-6])[）)]/);
  if (!volMatch) return null;
  const volMap = { '１': 1, '２': 2, '３': 3, '４': 4, '５': 5, '６': 6 };
  const volume = volMap[volMatch[1]] || parseInt(volMatch[1]);

  let version, variant;
  const isWhiteExam = productClass.includes('白卷');
  if (isWhiteExam) {
    if (product.includes('康版') || product.includes('[康版]')) {
      version = 'K';
      if (product.includes('A卷')) variant = 'a';
      else if (product.includes('B卷')) variant = 'b';
      else variant = '合計';
    } else if (product.includes('南版')) {
      version = 'N';
      variant = '合計';
    } else if (product.includes('翰版')) {
      version = 'H';
      if (product.includes('A卷')) variant = 'a';
      else if (product.includes('B卷')) variant = 'b';
      else variant = '合計';
    } else {
      return null;
    }
  } else if (productClass.includes('康卷')) {
    version = 'K';
    variant = productClass.includes('A卷') ? 'a' : 'b';
  } else if (productClass.includes('南卷')) {
    version = 'N';
    variant = '合計';
  } else if (productClass.includes('翰卷')) {
    version = 'H';
    variant = productClass.includes('A卷') ? 'a' : 'b';
  } else {
    return null;
  }

  let subject;
  if (product.includes('國文')) subject = '國文';
  else if (product.includes('英語') || product.includes('英文')) subject = '英文';
  else if (product.includes('數學')) subject = '數學';
  else if (product.includes('自然')) subject = '自然';
  else if (product.includes('地理')) subject = '地理';
  else if (product.includes('歷史')) subject = '歷史';
  else if (product.includes('公民')) subject = '公民';
  else return null;

  let grade;
  if (volume <= 2) grade = '一';
  else if (volume <= 4) grade = '三';
  else grade = '五';

  return { version, variant, subject, volume, grade, isWhiteExam };
}

// 1. Count parsed vs total records
let parsedCount = 0;
let unparsedCount = 0;
const unparsed = [];
for (const r of allRecords) {
  const parsed = parseExamProduct(r.product, r.productClass);
  if (parsed) {
    parsedCount++;
  } else {
    unparsedCount++;
    unparsed.push(`  ${r.product} [${r.productClass}] (sales: ${r.sales}, customer: ${r.customer}, qty: ${r.qty})`);
  }
}
console.log(`\n=== Record Parsing ===`);
console.log(`Total records: ${allRecords.length}`);
console.log(`Parsed: ${parsedCount}`);
console.log(`Unparsed: ${unparsedCount}`);
if (unparsed.length > 0) {
  console.log(`Unparsed records:`);
  for (const u of unparsed) console.log(u);
}

// 2. Verify totals per sales rep
console.log(`\n=== Totals per Sales Rep ===`);
for (const sales of targetSales) {
  const salesRecords = allRecords.filter(r => r.sales === sales);
  const totalQty = salesRecords.reduce((sum, r) => sum + r.qty, 0);
  
  // Sum only parsed records
  let parsedQty = 0;
  for (const r of salesRecords) {
    const parsed = parseExamProduct(r.product, r.productClass);
    if (parsed) parsedQty += r.qty;
  }
  
  console.log(`${sales}: JSON total=${totalQty}, parsed total=${parsedQty}, diff=${totalQty - parsedQty}`);
}

// 3. Verify each xlsx file
console.log(`\n=== Excel File Verification ===`);
const outputDir = path.join(__dirname, 'Output');

for (const sales of targetSales) {
  const filePath = path.join(outputDir, `${sales}_114上單冊考卷.xlsx`);
  if (!fs.existsSync(filePath)) {
    console.log(`✗ ${sales}: File not found!`);
    continue;
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  
  // Sum all data cells from each sheet (skip summary sheets)
  let excelTotal = 0;
  for (const ws of wb.worksheets) {
    if (ws.name === '客戶加總' || ws.name === '全部統計') continue;
    // The grand total row is the last row before the gap
    // Find it by looking for the "總數" label
    for (let r = 1; r <= ws.rowCount; r++) {
      const cellB = ws.getCell(r, 2).value;
      const cellC = ws.getCell(r, 3).value;
      // Grand total row has col3 = '總數' and is the last such row before summary
      if (cellC === '總數' || (typeof cellC === 'object' && cellC?.toString() === '總數')) {
        // Check if this is the grand total (not a sub-total)
        const cellA = ws.getCell(r, 1).value;
        // Grand total is the one with empty col1 AND col2
        if (!cellA && !cellB) {
          // Read the 實際出貨 column (col 27)
          const actualShipment = ws.getCell(r, 27).value;
          const val = typeof actualShipment === 'object' ? (actualShipment?.result || 0) : (actualShipment || 0);
          excelTotal += val;
        }
      }
    }
  }

  // Compare with JSON
  const salesRecords = allRecords.filter(r => r.sales === sales);
  let parsedQty = 0;
  for (const r of salesRecords) {
    const parsed = parseExamProduct(r.product, r.productClass);
    if (parsed) parsedQty += r.qty;
  }

  const match = excelTotal === parsedQty;
  console.log(`${match ? '✓' : '✗'} ${sales}: Excel total=${excelTotal}, JSON parsed total=${parsedQty} ${match ? '' : '(MISMATCH!)'}`);
}

console.log('\nVerification complete.');
