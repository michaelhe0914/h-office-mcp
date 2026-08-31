// Query h-office system for 其他 category product classes
// for 北區, 中區, and 南區 in the date range 114.4.26~114.9.25 (= 2025/4/26 ~ 2025/9/25)

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  const begin = "2025-4-26";
  const end = "2025-9-25";

  // The 5 "其他" category product classes
  const otherClasses = [
    "其他:補充類-國文",
    "其他:補充類-英語",
    "其他:12K喜悅",
    "其他:16K喜悅",
    "其他:其他543",
  ];

  // Verify login
  console.log("Verifying login status...");
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first.");
    return;
  }
  
  console.log("Logged in as:", parsedInitial.username);
  
  // Collect all records
  const allRecords = [];
  const targetZones = ["北區", "中區", "南區"];
  
  for (const pc of otherClasses) {
    console.log(`\nQuerying [${pc}] from ${begin} to ${end}...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);
    
    // Filter for the three regions
    const targetRecords = result.records.filter(r => targetZones.includes(r.zone));
    
    if (targetRecords.length > 0) {
      console.log(`  Found ${targetRecords.length} records in target zones.`);
      for (const r of targetRecords) {
        allRecords.push({
          productClass: pc,
          ...r
        });
      }
    } else {
      console.log(`  No records found in target zones.`);
    }
  }
  
  // Save all records to JSON
  fs.writeFileSync('query_results_other_543.json', JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\n=== Total records collected: ${allRecords.length} ===`);
  console.log(`Saved to query_results_other_543.json`);
  
  // Print summary by zone
  for (const zone of targetZones) {
    const zoneRecords = allRecords.filter(r => r.zone === zone);
    const totalQty = zoneRecords.reduce((sum, r) => sum + r.qty, 0);
    const totalRtn = zoneRecords.reduce((sum, r) => sum + r.rtn_qty, 0);
    console.log(`  ${zone}: ${zoneRecords.length} records, qty: ${totalQty}, rtn: ${totalRtn}, net: ${totalQty - totalRtn}`);
  }
  
  // Print summary by product class
  console.log("\n--- By Product Class ---");
  for (const pc of otherClasses) {
    const pcRecords = allRecords.filter(r => r.productClass === pc);
    const qty = pcRecords.reduce((sum, r) => sum + r.qty, 0);
    const rtn = pcRecords.reduce((sum, r) => sum + r.rtn_qty, 0);
    console.log(`  ${pc}: ${pcRecords.length} records, qty: ${qty}, rtn: ${rtn}, net: ${qty - rtn}`);
  }
}

run().catch(console.error);
