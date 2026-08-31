// Script to query 2025.7.01~9.25 北區 3900, 考前30天, UP+主題, 歷屆試題
import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  const begin = "2025-7-1";
  const end = "2025-9-25";

  const targetClasses = [
    "國中講義:3900題",
    "國中講義:考前30天",
    "國中講義:UP+主題",
    "國中講義:歷屆試題"
  ];

  console.log("Verifying login status...");
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);

  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please check session/cookies.");
    process.exit(1);
  }

  console.log(`Logged in as: ${parsedInitial.username}`);

  let allNorthRecords = [];

  for (const pc of targetClasses) {
    console.log(`Querying [${pc}] from ${begin} to ${end}...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);

    console.log(`  -> Fetched ${result.records.length} raw records for ${pc}`);
    
    // Filter for 北區 and attach productClass
    const northRecs = result.records
      .filter(r => r.zone === '北區')
      .map(r => ({ ...r, productClass: pc }));
      
    console.log(`  -> North Region records: ${northRecs.length}`);

    allNorthRecords = allNorthRecords.concat(northRecs);
  }

  console.log(`\nTotal North Region records fetched: ${allNorthRecords.length}`);

  // Save to raw JSON
  fs.writeFileSync('query_results_round2_review.json', JSON.stringify(allNorthRecords, null, 2), 'utf-8');
  console.log('Saved query_results_round2_review.json');
}

run().catch(err => {
  console.error(err);
  process.exit(1);
});
