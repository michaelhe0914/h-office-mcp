import { request, loadCookies } from './dist/client.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const res = await request(BASE_URL, "/");
  fs.writeFileSync('home_page.html', res.body, 'utf-8');
  console.log("Saved home_page.html");
}

run().catch(console.error);
