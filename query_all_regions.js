// Query h-office system for all 國中講義 single-volume product classes 
// for 北區, 中區, and 南區 in the date range 114.4.26~114.9.25 (= 2025/4/26 ~ 2025/9/25)
// Collect all records without filtering by specific sales reps

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  const begin = "2025-4-26";
  const end = "2025-9-25";

  // The product classes for 國中講義 single-volume (單冊)
  const singleVolumeClasses = [
    // 雙向
    "國中講義:雙向-康版(上)", "國中講義:雙向-康版(下)",
    "國中講義:雙向-南版(上)", "國中講義:雙向-南版(下)",
    "國中講義:雙向-翰版(上)", "國中講義:雙向-翰版(下)",
    // 735
    "國中講義:735-康版(上)", "國中講義:735-康版(下)",
    "國中講義:735-南版(上)", "國中講義:735-南版(下)",
    "國中講義:735-翰版(上)", "國中講義:735-翰版(下)",
    // 試題篇
    "國中講義:試題-康版(上)", "國中講義:試題-康版(下)",
    "國中講義:試題-南版(上)", "國中講義:試題-南版(下)",
    "國中講義:試題-翰版(上)", "國中講義:試題-翰版(下)",
    // 新講義
    "國中講義:新講義(上)", "國中講義:新講義(下)",
  ];

  // First, verify login
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
  
  for (const pc of singleVolumeClasses) {
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
  fs.writeFileSync('query_results_all_regions.json', JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\n=== Total records collected: ${allRecords.length} ===`);
  console.log(`Saved to query_results_all_regions.json`);
  
  // Print summary by zone
  for (const zone of targetZones) {
    const zoneRecords = allRecords.filter(r => r.zone === zone);
    const totalQty = zoneRecords.reduce((sum, r) => sum + r.qty, 0);
    console.log(`  ${zone}: ${zoneRecords.length} records, total qty: ${totalQty}`);
  }
}

run().catch(console.error);
