import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  const begin = "2025-4-26";
  const end = "2025-10-25";
  
  const classes = [
    "本土語:高中閩語-課本",
    "本土語:高中閩語-備課"
  ];

  console.log("Verifying login status...");
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first.");
    return;
  }
  
  console.log("Logged in as:", parsedInitial.username);
  
  const allRecords = [];
  
  for (const pc of classes) {
    console.log(`\nQuerying [${pc}] from ${begin} to ${end}...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);
    
    console.log(`  Found ${result.records.length} records.`);
    for (const r of result.records) {
      allRecords.push({
        productClass: pc,
        ...r
      });
    }
  }
  
  fs.writeFileSync('query_results_hs_tw.json', JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\n=== Total records collected: ${allRecords.length} ===`);
  console.log(`Saved to query_results_hs_tw.json`);
}

run().catch(console.error);
