// Query 中區 曹原菘 的 複習講義 + 新思維 本數
// 日期區間: 114.4.26~114.9.25 (= 2025-4-26 ~ 2025-9-25)

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  // 114.4.26 = 2025/4/26, 114.9.25 = 2025/9/25
  const begin = "2025-4-26";
  const end = "2025-9-25";
  const targetSales = "曹原菘";

  // Target product classes
  const targetClasses = [
    "國中講義:複習講義(不含5-6)",
    "國中講義:新思維(不含5-6)",
  ];

  // Check login first
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first.");
    return;
  }
  console.log("Logged in as:", parsedInitial.username);

  for (const pc of targetClasses) {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`查詢類別: ${pc}`);
    console.log(`時間範圍: ${begin} ~ ${end}`);
    console.log(`業務: ${targetSales}`);
    console.log(`${'='.repeat(60)}`);

    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);

    // Filter for 曹原菘 records
    const caoRecords = result.records.filter(r => r.sales === targetSales);

    if (caoRecords.length === 0) {
      console.log(`\n⚠️ 無出貨紀錄。`);
      continue;
    }

    // Group by customer
    const byCustomer = {};
    for (const r of caoRecords) {
      if (!byCustomer[r.customer]) byCustomer[r.customer] = [];
      byCustomer[r.customer].push(r);
    }

    let classTotalQty = 0;
    let classTotalRtn = 0;

    for (const [customer, records] of Object.entries(byCustomer).sort((a, b) => a[0].localeCompare(b[0]))) {
      console.log(`\n📌 經銷: ${customer}`);
      console.log(`  ${'產品'.padEnd(40)} | ${'訂量'.padStart(6)} | ${'退量'.padStart(6)}`);
      console.log(`  ${'-'.repeat(40)} | ${'-'.repeat(6)} | ${'-'.repeat(6)}`);
      
      let custQty = 0;
      let custRtn = 0;
      for (const r of records.sort((a, b) => a.product.localeCompare(b.product))) {
        const pName = r.product.padEnd(36);
        console.log(`  ${pName} | ${String(r.qty).padStart(6)} | ${String(r.rtn_qty).padStart(6)}`);
        custQty += r.qty;
        custRtn += r.rtn_qty;
      }
      console.log(`  👉 小計: 訂量 ${custQty} | 退量 ${custRtn} | 淨出 ${custQty - custRtn}`);
      classTotalQty += custQty;
      classTotalRtn += custRtn;
    }

    console.log(`\n✅ [${pc}] 合計: 訂量 ${classTotalQty} | 退量 ${classTotalRtn} | 淨出 ${classTotalQty - classTotalRtn}`);
    console.log(`   經銷數: ${Object.keys(byCustomer).length}`);
  }
}

run().catch(console.error);
