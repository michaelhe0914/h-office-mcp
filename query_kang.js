// Re-query for 康晉瑋 (correct name from system)
import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  const begin = "2025-4-26";
  const end = "2025-9-25";
  const targetSales = ["康晉瑋"]; // Correct name from system

  const singleVolumeClasses = [
    "國中講義:雙向-康版(上)", "國中講義:雙向-康版(下)",
    "國中講義:雙向-南版(上)", "國中講義:雙向-南版(下)",
    "國中講義:雙向-翰版(上)", "國中講義:雙向-翰版(下)",
    "國中講義:試題-康版(上)", "國中講義:試題-康版(下)",
    "國中講義:試題-南版(上)", "國中講義:試題-南版(下)",
    "國中講義:試題-翰版(上)", "國中講義:試題-翰版(下)",
    "國中講義:735-康版(上)", "國中講義:735-康版(下)",
    "國中講義:735-南版(上)", "國中講義:735-南版(下)",
    "國中講義:735-翰版(上)", "國中講義:735-翰版(下)",
    "國中講義:新講義(上)", "國中講義:新講義(下)",
  ];

  const allRecords = [];

  for (const pc of singleVolumeClasses) {
    console.log(`Querying [${pc}]...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);
    
    const targetRecords = result.records.filter(r => 
      r.zone === "北區" && targetSales.includes(r.sales)
    );
    
    if (targetRecords.length > 0) {
      console.log(`  Found ${targetRecords.length} records`);
      for (const r of targetRecords) {
        allRecords.push({ productClass: pc, ...r });
      }
    }
  }

  // Load existing data and add 康晉瑋 records
  const existingData = JSON.parse(fs.readFileSync('query_results_north_single_volume.json', 'utf-8'));
  const combined = [...existingData, ...allRecords];
  fs.writeFileSync('query_results_north_single_volume.json', JSON.stringify(combined, null, 2), 'utf-8');
  
  console.log(`\n康晉瑋: ${allRecords.length} records, total qty: ${allRecords.reduce((s, r) => s + r.qty, 0)}`);
  
  // Group by customer
  const byCustomer = {};
  for (const r of allRecords) {
    if (!byCustomer[r.customer]) byCustomer[r.customer] = [];
    byCustomer[r.customer].push(r);
  }
  for (const [customer, records] of Object.entries(byCustomer)) {
    const custQty = records.reduce((sum, r) => sum + r.qty, 0);
    console.log(`  ${customer}: ${records.length} products, qty: ${custQty}`);
  }
}

run().catch(console.error);
