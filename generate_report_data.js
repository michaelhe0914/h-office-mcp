import fs from 'fs';

// Read results from main query
const mainRecords = JSON.parse(fs.readFileSync('query_results_north_short.json', 'utf-8'));

// We also have the 5-6 data from the other query. Let's load the raw response of 新思維5-6 to extract it.
// Wait, we can just run a quick query to fetch 5-6 records and save them, or we can query it inside this script.
// Let's import client & parser to fetch the 5-6 records as well so that we have complete data.
import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const begin = "2025-4-26";
  const end = "2025-6-30";
  
  // Query 5-6
  const pc56 = "國中講義:新思維5-6";
  const params56 = { begin, end, product_class: pc56 };
  const res56 = await request(BASE_URL, "/sales", { params: params56 });
  const result56 = parseSalesPage(res56.body);
  const north56 = result56.records.filter(r => r.zone === '北區');

  const records56 = north56.map(r => ({
    productClass: pc56,
    zone: r.zone,
    sales: r.sales,
    customer: r.customer,
    product: r.product,
    qty: r.qty,
    rtn_qty: r.rtn_qty
  }));

  // Combine or keep separate
  // Let's summarize the standard 20 products (不含5-6) first
  console.log("### 國中複習講義 (不含5-6) 與 新思維 (不含5-6)");
  
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
