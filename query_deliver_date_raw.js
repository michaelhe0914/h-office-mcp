import { request, loadCookies } from './dist/client.js';
import * as fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  console.log("Fetching /deliver_date...");
  const res = await request(BASE_URL, "/deliver_date");
  console.log("Status Code:", res.statusCode);
  
  fs.writeFileSync('./deliver_date_raw.html', res.body, 'utf8');
  console.log("Saved raw HTML to deliver_date_raw.html");
  
  // Print some snippet
  console.log(res.body.substring(0, 1500));
}

run().catch(console.error);
