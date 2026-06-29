// Query ALL regions' 複習講義 + 新思維 data
// Date range: 114.4.26~114.9.25 (= 2025-4-26 ~ 2025-9-25)
// Save all records to JSON for XLSX generation

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  const begin = "2025-4-26";
  const end = "2025-9-25";

  const targetClasses = [
    "國中講義:複習講義(不含5-6)",
    "國中講義:新思維(不含5-6)",
  ];

  // Check login
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first.");
    process.exit(1);
  }
  console.log("Logged in as:", parsedInitial.username);

  const allRecords = [];

  for (const pc of targetClasses) {
    console.log(`\nQuerying [${pc}] from ${begin} to ${end}...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);

    console.log(`  Total records: ${result.records.length}`);
    
    // Keep ALL records from ALL zones
    for (const r of result.records) {
      allRecords.push({
        productClass: pc,
        zone: r.zone,
        sales: r.sales,
        customer: r.customer,
        product: r.product,
        qty: r.qty,
        rtn_qty: r.rtn_qty,
      });
    }
    
    // Print zone summary
    const byZone = {};
    for (const r of result.records) {
      if (!byZone[r.zone]) byZone[r.zone] = { qty: 0, rtn: 0, count: 0 };
      byZone[r.zone].qty += r.qty;
      byZone[r.zone].rtn += r.rtn_qty;
      byZone[r.zone].count += 1;
    }
    for (const [zone, data] of Object.entries(byZone)) {
      console.log(`  ${zone}: ${data.count} records, 訂=${data.qty}, 退=${data.rtn}`);
    }
  }

  // Save to JSON
  const outputPath = 'query_results_all_regions_review.json';
  fs.writeFileSync(outputPath, JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\nTotal records: ${allRecords.length}`);
  console.log(`Saved to ${outputPath}`);

  // Print overall summary
  const zones = [...new Set(allRecords.map(r => r.zone))].sort();
  const classes = [...new Set(allRecords.map(r => r.productClass))];
  console.log(`\nZones found: ${zones.join(', ')}`);
  console.log(`Product classes: ${classes.join(', ')}`);
  
  // Summary table
  console.log('\n=== Summary by Zone x Class ===');
  for (const zone of zones) {
    for (const pc of classes) {
      const recs = allRecords.filter(r => r.zone === zone && r.productClass === pc);
      const qty = recs.reduce((s, r) => s + r.qty, 0);
      const rtn = recs.reduce((s, r) => s + r.rtn_qty, 0);
      console.log(`  ${zone} | ${pc}: 訂=${qty} 退=${rtn} 淨=${qty - rtn}`);
    }
  }
}

run().catch(console.error);
