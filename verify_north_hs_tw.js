import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Canonical products
const TEXTBOOK_PRODUCTS = [
  '*高中職 閩南語文(１)',
  '*高中職 閩南語文(２)',
  '高中職 閩南語文(全一冊)上',
  '高中職 閩南語文(全一冊)下',
  '高中職 閩南語文(全一冊)乙版',
  '*高中職 閩南語文(全一冊)'
];

const PREP_PRODUCTS = [
  '*高中職 閩南語文(１)備課用書包'
];

async function run() {
  const jsonPath = path.join(__dirname, 'query_results_hs_tw.json');
  const xlsxPath = path.join(__dirname, 'Output', '高中台語出書統計表_北區.xlsx');

  if (!fs.existsSync(jsonPath)) {
    console.error(`❌ JSON data not found: ${jsonPath}`);
    process.exit(1);
  }
  if (!fs.existsSync(xlsxPath)) {
    console.error(`❌ Excel file not found: ${xlsxPath}`);
    process.exit(1);
  }

  console.log("Loading raw JSON data...");
  const rawRecords = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
  
  // Aggregate JSON totals for North
  const jsonCustomerTotals = {};
  const jsonNorthTotals = {};
  for (const prod of [...TEXTBOOK_PRODUCTS, ...PREP_PRODUCTS]) {
    jsonNorthTotals[prod] = { qty: 0, rtn: 0 };
  }

  for (const r of rawRecords) {
    if (r.zone !== '北區') continue;
    
    const cust = r.customer;
    if (!jsonCustomerTotals[cust]) {
      jsonCustomerTotals[cust] = { qty: 0, rtn: 0, prods: {} };
      for (const prod of [...TEXTBOOK_PRODUCTS, ...PREP_PRODUCTS]) {
        jsonCustomerTotals[cust].prods[prod] = { qty: 0, rtn: 0 };
      }
    }
    
    jsonCustomerTotals[cust].qty += r.qty;
    jsonCustomerTotals[cust].rtn += r.rtn_qty;
    
    if (jsonCustomerTotals[cust].prods[r.product]) {
      jsonCustomerTotals[cust].prods[r.product].qty += r.qty;
      jsonCustomerTotals[cust].prods[r.product].rtn += r.rtn_qty;
    }
    
    if (jsonNorthTotals[r.product]) {
      jsonNorthTotals[r.product].qty += r.qty;
      jsonNorthTotals[r.product].rtn += r.rtn_qty;
    }
  }

  console.log("Loading Excel file...");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  console.log(`Loaded. Worksheets: ${wb.worksheets.length}`);

  let errors = 0;

  // Helper to verify product rows in a sheet
  function verifyProductRows(ws, sheetLabel, expectedProds) {
    let sheetErrors = 0;
    
    // 1. Verify Textbooks (Rows 5 to 10)
    for (let i = 0; i < TEXTBOOK_PRODUCTS.length; i++) {
      const prod = TEXTBOOK_PRODUCTS[i];
      const rowNum = 5 + i;
      const row = ws.getRow(rowNum);
      
      const qtyVal = row.getCell(4).value || 0;
      const rtnVal = row.getCell(5).value || 0;
      const expected = expectedProds[prod] || { qty: 0, rtn: 0 };
      
      if (qtyVal !== expected.qty) {
        console.error(`  ❌ [${sheetLabel}] Row ${rowNum} (${prod}) Qty mismatch! Excel: ${qtyVal}, Expected: ${expected.qty}`);
        sheetErrors++;
      }
      if (rtnVal !== expected.rtn) {
        console.error(`  ❌ [${sheetLabel}] Row ${rowNum} (${prod}) Rtn mismatch! Excel: ${rtnVal}, Expected: ${expected.rtn}`);
        sheetErrors++;
      }
    }
    
    // 2. Verify Prep (Row 13)
    for (let i = 0; i < PREP_PRODUCTS.length; i++) {
      const prod = PREP_PRODUCTS[i];
      const rowNum = 13 + i;
      const row = ws.getRow(rowNum);
      
      const qtyVal = row.getCell(4).value || 0;
      const rtnVal = row.getCell(5).value || 0;
      const expected = expectedProds[prod] || { qty: 0, rtn: 0 };
      
      if (qtyVal !== expected.qty) {
        console.error(`  ❌ [${sheetLabel}] Row ${rowNum} (${prod}) Qty mismatch! Excel: ${qtyVal}, Expected: ${expected.qty}`);
        sheetErrors++;
      }
      if (rtnVal !== expected.rtn) {
        console.error(`  ❌ [${sheetLabel}] Row ${rowNum} (${prod}) Rtn mismatch! Excel: ${rtnVal}, Expected: ${expected.rtn}`);
        sheetErrors++;
      }
    }
    
    return sheetErrors;
  }

  // 1. Verify summary sheet
  const wsSummary = wb.getWorksheet('北區總表');
  if (!wsSummary) {
    console.error(`❌ Missing summary sheet: 北區總表`);
    errors++;
  } else {
    console.log("Verifying summary sheet [北區總表]...");
    const sheetErrors = verifyProductRows(wsSummary, '北區總表', jsonNorthTotals);
    errors += sheetErrors;
    if (sheetErrors === 0) {
      console.log(`  ✓ Summary Sheet [北區總表] matches JSON data.`);
    }
  }

  // 2. Verify individual customer sheets
  console.log(`\nVerifying ${Object.keys(jsonCustomerTotals).length} Customer Sheets for 北區...`);
  for (const [cust, jsonTot] of Object.entries(jsonCustomerTotals)) {
    const sheetName = cust.length > 31 ? cust.substring(0, 31) : cust;
    const ws = wb.getWorksheet(sheetName);
    
    if (!ws) {
      console.error(`❌ Missing sheet for customer: ${cust} (Sheet name: ${sheetName})`);
      errors++;
      continue;
    }
    
    const sheetErrors = verifyProductRows(ws, sheetName, jsonTot.prods);
    errors += sheetErrors;
  }

  if (errors === 0) {
    console.log("\n==================================================");
    console.log("💚 SUCCESS: All North Region verification checks passed!");
    console.log("==================================================");
  } else {
    console.log("\n==================================================");
    console.log(`❌ FAILED: Found ${errors} verification errors.`);
    console.log("==================================================");
    process.exit(1);
  }
}

run().catch(console.error);
