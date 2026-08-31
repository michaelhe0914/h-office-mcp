import { request, loadCookies } from './dist/client.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  const page = "/battlefield";
  console.log(`Querying ${page}...`);
  try {
    const res = await request(BASE_URL, page);
    fs.writeFileSync(`battlefield_page.html`, res.body, 'utf-8');
    console.log(`  Saved battlefield_page.html`);
  } catch (e) {
    console.error(`  Error querying ${page}:`, e);
  }
}

run().catch(console.error);
