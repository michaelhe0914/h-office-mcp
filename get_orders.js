import { request, loadCookies } from './dist/client.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const url = "/orders";
  console.log(`Querying ${url}...`);
  const res = await request(BASE_URL, url);
  fs.writeFileSync('orders_page.html', res.body, 'utf-8');
  console.log("Saved orders_page.html");
}

run().catch(console.error);
