import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Paths
const jsonPath = path.join(__dirname, 'query_results_north_review_exams.json');
const xlsxPath = path.join(__dirname, 'Output2', '114上_北區複習卷統計.xlsx');

async function run() {
  console.log("Loading files for verification (with asterisk-merged product names)...");
  
  // 1. Load JSON data
  const rawRecords = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
  console.log(`Loaded JSON: ${rawRecords.length} records`);

  // Aggregate quantities from JSON by [productClass, cleanProductName]
  const jsonTotals = {}; // `${productClass}::${cleanName}` -> { qty, rtn }
  const jsonCustomerTotals = {}; // customer -> productClass -> cleanName -> { qty, rtn }
  
  for (const r of rawRecords) {
    const cleanName = r.product.replace(/^\*/, '');
    const key = `${r.productClass}::${cleanName}`;
    if (!jsonTotals[key]) jsonTotals[key] = { qty: 0, rtn: 0 };
    jsonTotals[key].qty += r.qty;
    jsonTotals[key].rtn += r.rtn_qty;

    if (!jsonCustomerTotals[r.customer]) jsonCustomerTotals[r.customer] = {};
    if (!jsonCustomerTotals[r.customer][r.productClass]) jsonCustomerTotals[r.customer][r.productClass] = {};
    if (!jsonCustomerTotals[r.customer][r.productClass][cleanName]) {
      jsonCustomerTotals[r.customer][r.productClass][cleanName] = { qty: 0, rtn: 0 };
    }
    jsonCustomerTotals[r.customer][r.productClass][cleanName].qty += r.qty;
    jsonCustomerTotals[r.customer][r.productClass][cleanName].rtn += r.rtn_qty;
  }

  // 2. Load Excel Workbook
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);
  console.log(`Loaded XLSX. Total sheets: ${wb.worksheets.length}`);

  // 3. Verify sheets exist for all customers
  const sheetNames = wb.worksheets.map(ws => ws.name);
  console.log(`Summary sheet exists: ${sheetNames.includes('總表總量')}`);

  const rawCustomers = [...new Set(rawRecords.map(r => r.customer))];
  const missingSheets = [];
  for (const cust of rawCustomers) {
    let expectedSheetName = cust.replace(/[\\\/?*\[\]]/g, '');
    if (expectedSheetName.length > 31) expectedSheetName = expectedSheetName.substring(0, 31);
    if (!sheetNames.includes(expectedSheetName)) {
      missingSheets.push({ customer: cust, expectedSheetName });
    }
  }
  if (missingSheets.length > 0) {
    console.error("❌ ERROR: Missing sheets for customers:", missingSheets);
  } else {
    console.log("✓ Success: All customers have matching sheets in Excel.");
  }

  // 4. Verify Summary sheet totals against JSON totals
  const wsSummary = wb.getWorksheet('總表總量');
  let summaryErrors = 0;
  
  let rowCount = wsSummary.rowCount;
  console.log(`Summary sheet has ${rowCount} rows`);

  for (let r = 3; r <= rowCount; r++) { // Data starts at row 3
    const row = wsSummary.getRow(r);
    const indexVal = row.getCell(1).value;
    const classVal = row.getCell(2).value;
    const nameVal = row.getCell(3).value; // Already cleaned in Excel

    // Check if it's a data row (index is a number)
    if (typeof indexVal === 'number') {
      const qtyVal = row.getCell(4).value || 0;
      const rtnVal = row.getCell(5).value || 0;
      
      // Determine the full product class name
      let fullClass = "";
      if (classVal === "複習卷-A卷") fullClass = "國中考卷:複習卷-A卷";
      else if (classVal === "複習卷-B卷") fullClass = "國中考卷:複習卷-B卷";
      else if (classVal === "複習卷-其他") fullClass = "國中考卷:複習卷-其他";
      else if (classVal === "複習卷-新思維") fullClass = "國中考卷:複習卷-新思維";
      else if (classVal === "複習卷-半全冊") fullClass = "國中考卷:複習卷-半全冊";

      const jsonKey = `${fullClass}::${nameVal}`;
      const foundRecord = jsonTotals[jsonKey];

      if (!foundRecord) {
        if (qtyVal !== 0 || rtnVal !== 0) {
          console.error(`❌ ERROR: Row ${r} product [${nameVal}] under [${classVal}] has non-zero quantities (qty=${qtyVal}, rtn=${rtnVal}) but is NOT found in JSON!`);
          summaryErrors++;
        }
      } else {
        // Compare values
        if (qtyVal !== foundRecord.qty) {
          console.error(`❌ ERROR: Qty mismatch for ${jsonKey}. Excel: ${qtyVal}, JSON: ${foundRecord.qty}`);
          summaryErrors++;
        }
        if (rtnVal !== foundRecord.rtn) {
          console.error(`❌ ERROR: Return qty mismatch for ${jsonKey}. Excel: ${rtnVal}, JSON: ${foundRecord.rtn}`);
          summaryErrors++;
        }
      }
    }
  }

  if (summaryErrors === 0) {
    console.log("✓ Success: Summary sheet quantities match JSON raw data exactly!");
  } else {
    console.error(`❌ Total summary sheet mismatch errors: ${summaryErrors}`);
  }

  // 5. Cross-verify that sum of customer sheet quantities matches Summary sheet quantities row-by-row
  console.log("Verifying row-by-row customer sheet sum matches Summary sheet...");
  let rowMatches = 0;
  let rowMismatches = 0;

  for (let r = 3; r <= rowCount; r++) {
    const sumRow = wsSummary.getRow(r);
    const indexVal = sumRow.getCell(1).value;
    const nameVal = sumRow.getCell(3).value;

    if (typeof indexVal === 'number') {
      const sumQty = sumRow.getCell(4).value || 0;
      const sumRtn = sumRow.getCell(5).value || 0;

      let customersQtySum = 0;
      let customersRtnSum = 0;

      for (const sheetName of sheetNames) {
        if (sheetName === '總表總量') continue;
        const wsCust = wb.getWorksheet(sheetName);
        const custRow = wsCust.getRow(r);

        // Check if name aligns
        if (custRow.getCell(3).value !== nameVal) {
          console.error(`❌ Alignment error: Row ${r} on sheet [${sheetName}] name is [${custRow.getCell(3).value}], expected [${nameVal}]`);
          rowMismatches++;
          continue;
        }

        customersQtySum += (custRow.getCell(4).value || 0);
        customersRtnSum += (custRow.getCell(5).value || 0);
      }

      if (customersQtySum !== sumQty) {
        console.error(`❌ Qty mismatch on row ${r} [${nameVal}]. Summary sheet: ${sumQty}, sum of customers: ${customersQtySum}`);
        rowMismatches++;
      } else if (customersRtnSum !== sumRtn) {
        console.error(`❌ Return Qty mismatch on row ${r} [${nameVal}]. Summary sheet: ${sumRtn}, sum of customers: ${customersRtnSum}`);
        rowMismatches++;
      } else {
        rowMatches++;
      }
    }
  }

  console.log(`Verification results: ${rowMatches} product rows matched exactly. Mismatches: ${rowMismatches}`);
  if (rowMismatches === 0) {
    console.log("✓ Success: The sum of all customer sheets matches the Summary sheet exactly on every single row!");
  } else {
    console.error("❌ Mismatches found! Code needs correction.");
  }
}

run().catch(console.error);
