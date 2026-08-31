import { request, loadCookies } from './dist/client.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  // Query for 新時代書局 (040051)
  const url = "/deliver_date/040051";
  console.log(`Querying ${url}...`);
  const res = await request(BASE_URL, url);
  fs.writeFileSync('dealer_deliver_date.html', res.body, 'utf-8');
  console.log("Saved dealer_deliver_date.html");
}

run().catch(console.error);
