import { request, loadCookies } from './dist/client.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const url = "/orders/80005803";
  console.log(`Querying ${url}...`);
  const res = await request(BASE_URL, url);
  fs.writeFileSync('order_detail.html', res.body, 'utf-8');
  console.log("Saved order_detail.html");
}

run().catch(console.error);
