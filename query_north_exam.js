// Query h-office system for all 國中考卷 single-volume product classes
// for 北區 in the date range 114.4.26~114.9.25 (= 2025/4/26 ~ 2025/9/25)
// Sales reps: 何光傑、李敏豪、林智偉、康晉瑋、朱鵬學

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  // 114.4.26 = 2025/4/26, 114.9.25 = 2025/9/25
  const begin = "2025-4-26";
  const end = "2025-9-25";

  // Target sales reps (note: 康晉瑋 uses 瑋 in system)
  const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉瑋", "朱鵬學"];

  // 國中考卷 single-volume product classes (12 total)
  const targetProductClasses = [
    // 康卷-A卷
    "國中考卷:康卷-A卷(上)",
    "國中考卷:康卷-A卷(下)",
    // 康卷-B卷
    "國中考卷:康卷-B卷(上)",
    "國中考卷:康卷-B卷(下)",
    // 南卷
    "國中考卷:南卷(上)",
    "國中考卷:南卷(下)",
    // 翰卷-A卷
    "國中考卷:翰卷-A卷(上)",
    "國中考卷:翰卷-A卷(下)",
    // 翰卷-B卷
    "國中考卷:翰卷-B卷(上)",
    "國中考卷:翰卷-B卷(下)",
    // 8K白卷
    "國中考卷:8k單冊白卷(上)",
    "國中考卷:8k單冊白卷(下)",
  ];

  // Verify login
  console.log("Checking login status...");
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first.");
    return;
  }
  
  console.log("Logged in as:", parsedInitial.username);
  
  // Verify product classes exist
  const allClasses = parsedInitial.productClasses.map(c => c.name);
  console.log("\nVerifying target product classes...");
  for (const pc of targetProductClasses) {
    if (allClasses.includes(pc)) {
      console.log(`  ✓ ${pc}`);
    } else {
      console.log(`  ✗ ${pc} (NOT FOUND!)`);
    }
  }
  
  // Collect all records
  const allRecords = [];
  
  for (const pc of targetProductClasses) {
    console.log(`\nQuerying [${pc}] from ${begin} to ${end}...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);
    
    // Filter for 北區 and target sales reps
    const targetRecords = result.records.filter(r => 
      r.zone === "北區" && targetSales.includes(r.sales)
    );
    
    if (targetRecords.length > 0) {
      console.log(`  Found ${targetRecords.length} matching records`);
      for (const r of targetRecords) {
        allRecords.push({
          productClass: pc,
          ...r
        });
      }
    } else {
      console.log(`  No matching records for 北區 target sales reps.`);
    }
  }
  
  // Save all records to JSON
  fs.writeFileSync('query_results_north_exam.json', JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\n=== Total records collected: ${allRecords.length} ===`);
  console.log(`Saved to query_results_north_exam.json`);
  
  // Print summary by sales rep
  for (const sales of targetSales) {
    const salesRecords = allRecords.filter(r => r.sales === sales);
    const totalQty = salesRecords.reduce((sum, r) => sum + r.qty, 0);
    console.log(`\n${sales}: ${salesRecords.length} records, total qty: ${totalQty}`);
    
    // Group by customer
    const byCustomer = {};
    for (const r of salesRecords) {
      if (!byCustomer[r.customer]) byCustomer[r.customer] = [];
      byCustomer[r.customer].push(r);
    }
    for (const [customer, records] of Object.entries(byCustomer)) {
      const custQty = records.reduce((sum, r) => sum + r.qty, 0);
      console.log(`  ${customer}: ${records.length} products, qty: ${custQty}`);
    }
  }
  
  // Print sample product names for analysis
  console.log("\n=== Sample product names by productClass ===");
  const byPC = {};
  for (const r of allRecords) {
    if (!byPC[r.productClass]) byPC[r.productClass] = new Set();
    byPC[r.productClass].add(r.product);
  }
  for (const [pc, products] of Object.entries(byPC)) {
    console.log(`\n[${pc}]:`);
    for (const p of [...products].slice(0, 5)) {
      console.log(`  ${p}`);
    }
    if (products.size > 5) console.log(`  ... and ${products.size - 5} more`);
  }
}

run().catch(console.error);
