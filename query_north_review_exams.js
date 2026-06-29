import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const begin = "2025-4-26";
  const end = "2025-9-25";

  const targetClasses = [
    "國中考卷:複習卷-A卷",
    "國中考卷:複習卷-B卷",
    "國中考卷:複習卷-其他",
    "國中考卷:複習卷-新思維",
    "國中考卷:複習卷-半全冊"
  ];

  // Verify login status
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please run npm run login first.");
    process.exit(1);
  }
  console.log("Logged in as:", parsedInitial.username);

  const allRecords = [];

  for (const pc of targetClasses) {
    console.log(`Querying [${pc}] from ${begin} to ${end}...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);

    console.log(`  Total records returned: ${result.records.length}`);
    
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

  // Save to JSON file
  const outputPath = path.join(__dirname, 'query_results_north_review_exams.json');
  fs.writeFileSync(outputPath, JSON.stringify(allRecords, null, 2), 'utf-8');
  console.log(`\nSaved ${allRecords.length} Northern region records to ${outputPath}`);

  // Print some stats
  const customers = [...new Set(allRecords.map(r => r.customer))].sort();
  console.log(`Unique customers (${customers.length}):`, customers.join(', '));

  const salesReps = [...new Set(allRecords.map(r => r.sales))].sort();
  console.log(`Sales reps (${salesReps.length}):`, salesReps.join(', '));
}

run().catch(console.error);
