// Check if 康晉偉 exists as a sales rep in the system
import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  // Query just one product class to see all sales reps
  const params = { 
    begin: "2025-4-26", 
    end: "2025-9-25", 
    product_class: "國中講義:雙向-康版(上)" 
  };
  const res = await request(BASE_URL, "/sales", { params });
  const result = parseSalesPage(res.body);
  
  // Get unique sales reps from 北區
  const northReps = new Set();
  const allSalesReps = new Set();
  for (const r of result.records) {
    allSalesReps.add(`${r.zone}|${r.sales}`);
    if (r.zone === "北區") {
      northReps.add(r.sales);
    }
  }
  
  console.log("北區 Sales Reps:");
  [...northReps].sort().forEach(s => console.log(`  ${s}`));
  
  console.log("\nAll Sales Reps:");
  [...allSalesReps].sort().forEach(s => console.log(`  ${s}`));
  
  // Check if 康晉偉 appears anywhere
  const kangRecords = result.records.filter(r => r.sales.includes("康"));
  console.log(`\nRecords with "康" in sales rep: ${kangRecords.length}`);
  kangRecords.forEach(r => console.log(`  ${r.zone} | ${r.sales} | ${r.customer}`));
}

run().catch(console.error);
