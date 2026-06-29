// Query Northern region's 複習講義 and 新思維 data for Year 115 (2026)
// Date range: 115.4.26~115.6.30 (= 2026-4-26 ~ 2026-6-30)

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  const begin = "2026-4-26";
  const end = "2026-6-30";

  const targetClasses = [
    "國中講義:複習講義(不含5-6)",
    "國中講義:新思維(不含5-6)",
    "國中講義:新思維5-6",
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
  const outputPath = 'query_results_north_115.json';
  fs.writeFileSync(outputPath, JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\nSaved ${allRecords.length} Northern region records to ${outputPath}`);

  // We group by category
  console.log("\n### 國中複習講義 (不含5-6) 與 新思維 (不含5-6)");
  
  const mainRecords = allRecords.filter(r => r.productClass !== "國中講義:新思維5-6");
  const records56 = allRecords.filter(r => r.productClass === "國中講義:新思維5-6");

  const categories = {
    "雙向複習系列": mainRecords.filter(r => r.product.includes("雙向")),
    "735輕鬆讀系列": mainRecords.filter(r => r.product.includes("735")),
    "主題讚系列": mainRecords.filter(r => r.product.includes("主題讚")),
    "新思維系列 (不含5-6)": mainRecords.filter(r => r.productClass.includes("新思維") && !r.product.includes("5-6"))
  };

  let grandTotalQty = 0;
  let grandTotalRtn = 0;
  let grandTotalNet = 0;

  for (const [catName, recs] of Object.entries(categories)) {
    console.log(`\n#### ${catName}`);
    console.log(`| 產品名稱 | 訂量 | 退量 | 淨出貨 | 退書率 |`);
    console.log(`| :--- | ---: | ---: | ---: | ---: |`);
    
    // Group by product
    const prodMap = {};
    for (const r of recs) {
      if (!prodMap[r.product]) prodMap[r.product] = { qty: 0, rtn: 0 };
      prodMap[r.product].qty += r.qty;
      prodMap[r.product].rtn += r.rtn_qty;
    }
    
    let catQty = 0;
    let catRtn = 0;
    
    for (const [prodName, data] of Object.entries(prodMap).sort((a,b)=>a[0].localeCompare(b[0]))) {
      const net = data.qty - data.rtn;
      const rate = data.qty > 0 ? ((data.rtn / (data.qty + data.rtn)) * 100).toFixed(1) + '%' : '0.0%';
      console.log(`| ${prodName} | ${data.qty.toLocaleString()} | ${data.rtn.toLocaleString()} | **${net.toLocaleString()}** | ${rate} |`);
      catQty += data.qty;
      catRtn += data.rtn;
    }
    
    const catNet = catQty - catRtn;
    const catRate = catQty > 0 ? ((catRtn / (catQty + catRtn)) * 100).toFixed(1) + '%' : '0.0%';
    console.log(`| **${catName} 小計** | **${catQty.toLocaleString()}** | **${catRtn.toLocaleString()}** | **${catNet.toLocaleString()}** | **${catRate}** |`);
    
    grandTotalQty += catQty;
    grandTotalRtn += catRtn;
    grandTotalNet += catNet;
  }

  const grandRate = grandTotalQty > 0 ? ((grandTotalRtn / (grandTotalQty + grandTotalRtn)) * 100).toFixed(1) + '%' : '0.0%';
  console.log(`\n| **總計 (不含5-6)** | **${grandTotalQty.toLocaleString()}** | **${grandTotalRtn.toLocaleString()}** | **${grandTotalNet.toLocaleString()}** | **${grandRate}** |`);

  // Now output the 5-6 series
  console.log(`\n### 新思維複習系列 (5-6冊)`);
  console.log(`| 產品名稱 | 訂量 | 退量 | 淨出貨 | 退書率 |`);
  console.log(`| :--- | ---: | ---: | ---: | ---: |`);
  
  const prodMap56 = {};
  for (const r of records56) {
    if (!prodMap56[r.product]) prodMap56[r.product] = { qty: 0, rtn: 0 };
    prodMap56[r.product].qty += r.qty;
    prodMap56[r.product].rtn += r.rtn_qty;
  }
  
  let totalQty56 = 0;
  let totalRtn56 = 0;
  for (const [prodName, data] of Object.entries(prodMap56).sort((a,b)=>a[0].localeCompare(b[0]))) {
    const net = data.qty - data.rtn;
    const rate = data.qty > 0 ? ((data.rtn / (data.qty + data.rtn)) * 100).toFixed(1) + '%' : '0.0%';
    console.log(`| ${prodName} | ${data.qty.toLocaleString()} | ${data.rtn.toLocaleString()} | **${net.toLocaleString()}** | ${rate} |`);
    totalQty56 += data.qty;
    totalRtn56 += data.rtn;
  }
  const net56 = totalQty56 - totalRtn56;
  const rate56 = totalQty56 > 0 ? ((totalRtn56 / (totalQty56 + totalRtn56)) * 100).toFixed(1) + '%' : '0.0%';
  console.log(`| **新思維5-6 小計** | **${totalQty56.toLocaleString()}** | **${totalRtn56.toLocaleString()}** | **${net56.toLocaleString()}** | **${rate56}** |`);
}

run().catch(console.error);
