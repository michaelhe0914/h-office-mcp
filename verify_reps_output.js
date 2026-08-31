import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Paths
const jsonPath = path.join(__dirname, 'query_results_north_review_exams.json');
const output2Dir = path.join(__dirname, 'Output2');

const reps = ["何光傑", "康晉瑋", "朱鵬學", "李敏豪", "林智偉"];

async function run() {
  console.log("Loading raw JSON data...");
  const rawRecords = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
  
  // Aggregate JSON data by sales rep, customer, product class, and clean product name
  // rep -> productClass::productName -> { qty, rtn }
  const jsonRepTotals = {};
  
  for (const r of rawRecords) {
    const cleanName = r.product.replace(/^\*/, '');
    const rep = r.sales || '未指定業務';
    const key = `${r.productClass}::${cleanName}`;
    
    if (!jsonRepTotals[rep]) jsonRepTotals[rep] = {};
    if (!jsonRepTotals[rep][key]) jsonRepTotals[rep][key] = { qty: 0, rtn: 0 };
    
    jsonRepTotals[rep][key].qty += r.qty;
    jsonRepTotals[rep][key].rtn += r.rtn_qty;
  }

  let totalRepsVerified = 0;
  let overallMismatches = 0;

  for (const rep of reps) {
    const filename = `${rep}_114上_北區複習卷統計.xlsx`;
    const xlsxPath = path.join(output2Dir, filename);
    
    if (!fs.existsSync(xlsxPath)) {
      console.error(`❌ ERROR: Excel file not found for rep [${rep}]: ${xlsxPath}`);
      overallMismatches++;
      continue;
    }

    console.log(`\nVerifying file: ${filename}...`);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(xlsxPath);
    console.log(`  Loaded. Sheets: ${wb.worksheets.length}`);

    const sheetNames = wb.worksheets.map(ws => ws.name);
    if (!sheetNames.includes('總表總量')) {
      console.error(`  ❌ ERROR: Missing '總表總量' sheet in ${filename}`);
      overallMismatches++;
      continue;
    }

    // 1. Verify summary sheet totals against JSON totals for this rep
    const wsSummary = wb.getWorksheet('總表總量');
    const repJsonData = jsonRepTotals[rep] || {};
    let repSummaryErrors = 0;
    
    const rowCount = wsSummary.rowCount;

    for (let r = 3; r <= rowCount; r++) {
      const row = wsSummary.getRow(r);
      const indexVal = row.getCell(1).value;
      const classVal = row.getCell(2).value;
      const nameVal = row.getCell(3).value;

      if (typeof indexVal === 'number') {
        const qtyVal = row.getCell(4).value || 0;
        const rtnVal = row.getCell(5).value || 0;

        let fullClass = "";
        if (classVal === "複習卷-A卷") fullClass = "國中考卷:複習卷-A卷";
        else if (classVal === "複習卷-B卷") fullClass = "國中考卷:複習卷-B卷";
        else if (classVal === "複習卷-其他") fullClass = "國中考卷:複習卷-其他";
        else if (classVal === "複習卷-新思維") fullClass = "國中考卷:複習卷-新思維";
        else if (classVal === "複習卷-半全冊") fullClass = "國中考卷:複習卷-半全冊";

        const key = `${fullClass}::${nameVal}`;
        const found = repJsonData[key] || { qty: 0, rtn: 0 };

        if (qtyVal !== found.qty) {
          console.error(`  ❌ Qty mismatch for [${key}]. Excel: ${qtyVal}, JSON: ${found.qty}`);
          repSummaryErrors++;
        }
        if (rtnVal !== found.rtn) {
          console.error(`  ❌ Return qty mismatch for [${key}]. Excel: ${rtnVal}, JSON: ${found.rtn}`);
          repSummaryErrors++;
        }
      }
    }

    if (repSummaryErrors === 0) {
      console.log(`  ✓ Success: Summary sheet matches JSON data for ${rep}`);
    } else {
      console.error(`  ❌ Total summary mismatches for ${rep}: ${repSummaryErrors}`);
      overallMismatches += repSummaryErrors;
    }

    // 2. Verify row-by-row customer sheet sum matches Summary sheet
    let repRowMatches = 0;
    let repRowMismatches = 0;

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

          if (custRow.getCell(3).value !== nameVal) {
            console.error(`  ❌ Alignment error row ${r} on sheet [${sheetName}]: [${custRow.getCell(3).value}], expected [${nameVal}]`);
            repRowMismatches++;
            continue;
          }

          customersQtySum += (custRow.getCell(4).value || 0);
          customersRtnSum += (custRow.getCell(5).value || 0);
        }

        if (customersQtySum !== sumQty) {
          console.error(`  ❌ Row ${r} [${nameVal}] qty sum mismatch. Summary: ${sumQty}, customers: ${customersQtySum}`);
          repRowMismatches++;
        } else if (customersRtnSum !== sumRtn) {
          console.error(`  ❌ Row ${r} [${nameVal}] return sum mismatch. Summary: ${sumRtn}, customers: ${customersRtnSum}`);
          repRowMismatches++;
        } else {
          repRowMatches++;
        }
      }
    }

    if (repRowMismatches === 0) {
      console.log(`  ✓ Success: Sum of customer sheets matches summary sheet exactly for ${rep}`);
      totalRepsVerified++;
    } else {
      console.error(`  ❌ Row mismatches for ${rep}: ${repRowMismatches}`);
      overallMismatches += repRowMismatches;
    }
  }

  // 3. Verify partitioning: sum of rep files matches grand overall file
  console.log("\nVerifying partition completeness: Sum of reps' summaries matches grand overall summary...");
  const grandWb = new ExcelJS.Workbook();
  const grandPath = path.join(output2Dir, '114上_北區複習卷統計.xlsx');
  await grandWb.xlsx.readFile(grandPath);
  const grandSummaryWs = grandWb.getWorksheet('總表總量');
  const grandRowCount = grandSummaryWs.rowCount;

  // Load all rep summaries in memory
  const repWorkbooks = {};
  for (const rep of reps) {
    const filename = `${rep}_114上_北區複習卷統計.xlsx`;
    const xlsxPath = path.join(output2Dir, filename);
    const repWb = new ExcelJS.Workbook();
    await repWb.xlsx.readFile(xlsxPath);
    repWorkbooks[rep] = repWb.getWorksheet('總表總量');
  }

  let partitionMatches = 0;
  let partitionMismatches = 0;

  for (let r = 3; r <= grandRowCount; r++) {
    const grandRow = grandSummaryWs.getRow(r);
    const indexVal = grandRow.getCell(1).value;
    const nameVal = grandRow.getCell(3).value;

    if (typeof indexVal === 'number') {
      const grandQty = grandRow.getCell(4).value || 0;
      const grandRtn = grandRow.getCell(5).value || 0;

      let repsQtySum = 0;
      let repsRtnSum = 0;

      for (const rep of reps) {
        const repRow = repWorkbooks[rep].getRow(r);
        
        // Safety check name alignment
        if (repRow.getCell(3).value !== nameVal) {
          console.error(`❌ Alignment mismatch on partition row ${r} for rep ${rep}: [${repRow.getCell(3).value}], expected [${nameVal}]`);
          partitionMismatches++;
          continue;
        }

        repsQtySum += (repRow.getCell(4).value || 0);
        repsRtnSum += (repRow.getCell(5).value || 0);
      }

      if (repsQtySum !== grandQty) {
        console.error(`❌ Partition Qty mismatch on row ${r} [${nameVal}]. Grand overall: ${grandQty}, sum of reps: ${repsQtySum}`);
        partitionMismatches++;
      } else if (repsRtnSum !== grandRtn) {
        console.error(`❌ Partition Return Qty mismatch on row ${r} [${nameVal}]. Grand overall: ${grandRtn}, sum of reps: ${repsRtnSum}`);
        partitionMismatches++;
      } else {
        partitionMatches++;
      }
    }
  }

  console.log(`Partition verification: ${partitionMatches} rows verified. Mismatches: ${partitionMismatches}`);
  
  if (overallMismatches === 0 && partitionMismatches === 0) {
    console.log("\n🏆 ALL VERIFICATIONS PASSED SUCCESSFULLY! The per-sales-rep files and the overall file are 100% mathematically consistent!");
  } else {
    console.error(`\n❌ VERIFICATION FAILED! Total errors: ${overallMismatches + partitionMismatches}`);
  }
}

run().catch(console.error);
