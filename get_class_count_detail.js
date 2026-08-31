import { request, loadCookies } from './dist/client.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  const pages = [
    "/class_count_detail",
    "/class_count_all_detail" // Guessing there might be a detail page for class_count_all too
  ];
  
  for (const page of pages) {
    console.log(`Querying ${page}...`);
    try {
      const res = await request(BASE_URL, page);
      fs.writeFileSync(`${page.replace('/', '')}_page.html`, res.body, 'utf-8');
      console.log(`  Saved ${page.replace('/', '')}_page.html`);
    } catch (e) {
      console.error(`  Error querying ${page}:`, e);
    }
  }
}

run().catch(console.error);
