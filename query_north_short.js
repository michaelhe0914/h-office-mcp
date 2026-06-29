// Query Northern region's 複習講義 and 新思維 data
// Date range: 114.4.26~114.6.30 (= 2025-4-26 ~ 2025-6-30)
// Print summary directly.

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  const begin = "2025-4-26";
  const end = "2025-6-30";

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
    
    // Filter records for 北區
    const northRecords = result.records.filter(r => r.zone === '北區');
    console.log(`  Northern region records: ${northRecords.length}`);

    for (const r of northRecords) {
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
  }

  // Save to JSON
  const outputPath = 'query_results_north_short.json';
  fs.writeFileSync(outputPath, JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\nSaved ${allRecords.length} Northern region records to ${outputPath}`);

  // Summary by product
  const summary = {};
  for (const r of allRecords) {
    const key = `${r.productClass} | ${r.product}`;
    if (!summary[key]) {
      summary[key] = {
        productClass: r.productClass,
        product: r.product,
        qty: 0,
        rtn_qty: 0,
        net_qty: 0
      };
    }
    summary[key].qty += r.qty;
    summary[key].rtn_qty += r.rtn_qty;
    summary[key].net_qty += (r.qty - r.rtn_qty);
  }

  console.log('\n=== Summary by Product ===');
  const sortedKeys = Object.keys(summary).sort();
  for (const key of sortedKeys) {
    const s = summary[key];
    console.log(`${s.product}: 訂量=${s.qty}, 退量=${s.rtn_qty}, 淨出貨=${s.net_qty}`);
  }

  // Summary by Class
  console.log('\n=== Summary by Product Class ===');
  const classSummary = {};
  for (const r of allRecords) {
    if (!classSummary[r.productClass]) {
      classSummary[r.productClass] = { qty: 0, rtn_qty: 0, net_qty: 0 };
    }
    classSummary[r.productClass].qty += r.qty;
    classSummary[r.productClass].rtn_qty += r.rtn_qty;
    classSummary[r.productClass].net_qty += (r.qty - r.rtn_qty);
  }
  for (const [pc, s] of Object.entries(classSummary)) {
    console.log(`${pc}: 訂量=${s.qty}, 退量=${s.rtn_qty}, 淨出貨=${s.net_qty}`);
  }

  // Total
  const totalQty = allRecords.reduce((sum, r) => sum + r.qty, 0);
  const totalRtn = allRecords.reduce((sum, r) => sum + r.rtn_qty, 0);
  console.log(`\n總計: 訂量=${totalQty}, 退量=${totalRtn}, 淨出貨=${totalQty - totalRtn}`);
}

run().catch(console.error);
