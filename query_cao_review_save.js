// Query 中區 曹原菘 的 複習講義 + 新思維 完整資料
// 日期區間: 114.4.26~114.9.25 (= 2025-4-26 ~ 2025-9-25)
// 儲存為 JSON 供 XLSX 產生器使用

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  const begin = "2025-4-26";
  const end = "2025-9-25";
  const targetSales = "曹原菘";

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

    const caoRecords = result.records.filter(r => r.sales === targetSales);
    console.log(`  Found ${caoRecords.length} records for ${targetSales}`);

    for (const r of caoRecords) {
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
  const outputPath = 'query_results_cao_review_xinswei.json';
  fs.writeFileSync(outputPath, JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\nTotal records: ${allRecords.length}`);
  console.log(`Saved to ${outputPath}`);

  // Print summary
  const byClass = {};
  for (const r of allRecords) {
    if (!byClass[r.productClass]) byClass[r.productClass] = { qty: 0, rtn: 0, customers: new Set() };
    byClass[r.productClass].qty += r.qty;
    byClass[r.productClass].rtn += r.rtn_qty;
    byClass[r.productClass].customers.add(r.customer);
  }
  for (const [pc, data] of Object.entries(byClass)) {
    console.log(`\n[${pc}]`);
    console.log(`  訂量: ${data.qty}, 退量: ${data.rtn}, 淨出: ${data.qty - data.rtn}`);
    console.log(`  經銷數: ${data.customers.size}`);
  }
}

run().catch(console.error);
