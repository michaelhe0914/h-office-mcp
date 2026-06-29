// Script to query 國中講義:考前30天 for 2025.9/26 ~ 2026.5/25
// for 北區, 中區, 南區 and generate summary

import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  const begin = "2024-9-26";
  const end = "2025-5-25";
  const pc = "國中講義:考前30天";

  console.log("Verifying login status...");
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first using `npm run login` or manual cookie update.");
    return;
  }
  
  console.log("Logged in as:", parsedInitial.username);
  console.log(`Querying [${pc}] from ${begin} to ${end}...`);
  
  const params = { begin, end, product_class: pc };
  const res = await request(BASE_URL, "/sales", { params });
  const result = parseSalesPage(res.body);
  
  console.log(`Total raw records fetched: ${result.records.length}`);
  
  // Filter for North, Central, and South zones
  const targetZones = ["北區", "中區", "南區"];
  const filteredRecords = result.records.filter(r => targetZones.includes(r.zone));
  console.log(`Records in target zones (北區, 中區, 南區): ${filteredRecords.length}`);
  
  // Save filtered records to JSON for backup/inspection
  fs.writeFileSync('query_results_exam_30.json', JSON.stringify(filteredRecords, null, 2), 'utf-8');
  
  // Group by product
  // For each product, track:
  // - 北區: { qty, rtn_qty }
  // - 中區: { qty, rtn_qty }
  // - 南區: { qty, rtn_qty }
  const productSummary = {};
  
  for (const r of filteredRecords) {
    if (!productSummary[r.product]) {
      productSummary[r.product] = {
        "北區": { qty: 0, rtn_qty: 0 },
        "中區": { qty: 0, rtn_qty: 0 },
        "南區": { qty: 0, rtn_qty: 0 }
      };
    }
    productSummary[r.product][r.zone].qty += r.qty;
    productSummary[r.product][r.zone].rtn_qty += r.rtn_qty;
  }
  
  // Generate and print summary
  console.log("\n=== 考前30天 各區域品項統計 ===");
  
  // Let's sort products alphabetically or by some ordering if any
  const sortedProducts = Object.keys(productSummary).sort();
  
  const tableData = sortedProducts.map(prod => {
    const data = productSummary[prod];
    const nQty = data["北區"].qty;
    const nRtn = data["北區"].rtn_qty;
    const nNet = nQty - nRtn;
    
    const cQty = data["中區"].qty;
    const cRtn = data["中區"].rtn_qty;
    const cNet = cQty - cRtn;
    
    const sQty = data["南區"].qty;
    const sRtn = data["南區"].rtn_qty;
    const sNet = sQty - sRtn;
    
    const totalQty = nQty + cQty + sQty;
    const totalRtn = nRtn + cRtn + sRtn;
    const totalNet = totalQty - totalRtn;
    
    return {
      product: prod,
      north: { qty: nQty, rtn: nRtn, net: nNet },
      central: { qty: cQty, rtn: cRtn, net: cNet },
      south: { qty: sQty, rtn: sRtn, net: sNet },
      total: { qty: totalQty, rtn: totalRtn, net: totalNet }
    };
  });
  
  // Save formatted table data
  fs.writeFileSync('query_summary_exam_30.json', JSON.stringify(tableData, null, 2), 'utf-8');
  
  // Output Markdown Table
  let markdown = `| 品項 (Product) | 北區訂量 | 北區退量 | 北區淨量 | 中區訂量 | 中區退量 | 中區淨量 | 南區訂量 | 南區退量 | 南區淨量 | 總訂量 | 總退量 | 總淨量 |\n`;
  markdown += `| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |\n`;
  
  let grandNorthQty = 0, grandNorthRtn = 0, grandNorthNet = 0;
  let grandCentralQty = 0, grandCentralRtn = 0, grandCentralNet = 0;
  let grandSouthQty = 0, grandSouthRtn = 0, grandSouthNet = 0;
  let grandTotalQty = 0, grandTotalRtn = 0, grandTotalNet = 0;
  
  for (const row of tableData) {
    markdown += `| ${row.product} | ${row.north.qty} | ${row.north.rtn} | **${row.north.net}** | ${row.central.qty} | ${row.central.rtn} | **${row.central.net}** | ${row.south.qty} | ${row.south.rtn} | **${row.south.net}** | ${row.total.qty} | ${row.total.rtn} | **${row.total.net}** |\n`;
    
    grandNorthQty += row.north.qty;
    grandNorthRtn += row.north.rtn;
    grandNorthNet += row.north.net;
    
    grandCentralQty += row.central.qty;
    grandCentralRtn += row.central.rtn;
    grandCentralNet += row.central.net;
    
    grandSouthQty += row.south.qty;
    grandSouthRtn += row.south.rtn;
    grandSouthNet += row.south.net;
    
    grandTotalQty += row.total.qty;
    grandTotalRtn += row.total.rtn;
    grandTotalNet += row.total.net;
  }
  
  markdown += `| **加總 (Total)** | **${grandNorthQty}** | **${grandNorthRtn}** | **${grandNorthNet}** | **${grandCentralQty}** | **${grandCentralRtn}** | **${grandCentralNet}** | **${grandSouthQty}** | **${grandSouthRtn}** | **${grandSouthNet}** | **${grandTotalQty}** | **${grandTotalRtn}** | **${grandTotalNet}** |\n`;
  
  console.log(markdown);
  fs.writeFileSync('query_markdown_exam_30.md', markdown, 'utf-8');
}

run().catch(console.error);
