// First step: query h-office system for all 國中講義 single-volume product classes 
// for 北區 in the date range 114.4.26~114.9.25 (= 2025/4/26 ~ 2025/9/25)
// Collect all records for the 5 sales reps: 何光傑、李敏豪、林智偉、康晉偉、朱鵬學

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  // 114.4.26 = 2025/4/26, 114.9.25 = 2025/9/25
  const begin = "2025-4-26";
  const end = "2025-9-25";

  // Target sales reps
  const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉偉", "朱鵬學"];

  // The product classes for 國中講義 single-volume (單冊)
  // From the _product_classes variable in the sales page:
  const targetProductClasses = [
    // 雙向
    "國中講義:雙向-康版(上)",
    "國中講義:雙向-康版(下)",
    "國中講義:雙向-南版(上)", 
    "國中講義:雙向-南版(下)",
    "國中講義:雙向-翰版(上)",
    "國中講義:雙向-翰版(下)",
    // 735
    "國中講義:735-康版(上)",
    "國中講義:735-康版(下)",
    "國中講義:735-南版(上)",
    "國中講義:735-南版(下)",
    "國中講義:735-翰版(上)",
    "國中講義:735-翰版(下)",
    // 試題篇
    "國中講義:試題-康版(上)",
    "國中講義:試題-康版(下)",
    "國中講義:試題-南版(上)",
    "國中講義:試題-南版(下)",
    "國中講義:試題-翰版(上)",
    "國中講義:試題-翰版(下)",
    // 新講義
    "國中講義:新講義(上)",
    "國中講義:新講義(下)",
  ];

  // First, get the actual product classes from the system
  console.log("Fetching available product classes...");
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first.");
    return;
  }
  
  console.log("Logged in as:", parsedInitial.username);
  
  // Find matching product classes
  const allClasses = parsedInitial.productClasses.map(c => c.name);
  console.log("\nAll available product classes related to 國中講義 單冊:");
  const singleVolumeClasses = allClasses.filter(name => 
    name.includes("國中講義:雙向") || 
    name.includes("國中講義:735") || 
    name.includes("國中講義:試題") ||
    name.includes("國中講義:新講義")
  );
  console.log(singleVolumeClasses);
  
  // Collect all records
  const allRecords = [];
  
  for (const pc of singleVolumeClasses) {
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
  
  // Save all records to JSON for further processing
  fs.writeFileSync('query_results_north_single_volume.json', JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\n=== Total records collected: ${allRecords.length} ===`);
  console.log(`Saved to query_results_north_single_volume.json`);
  
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
}

run().catch(console.error);
