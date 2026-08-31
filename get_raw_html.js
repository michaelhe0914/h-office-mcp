import { request, loadCookies } from './dist/client.js';
import fs from 'fs';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const params = {
    begin: "2025-4-26",
    end: "2025-10-25",
    product_class: "本土語:高中閩語-課本"
  };
  const res = await request(BASE_URL, "/sales", { params });
  fs.writeFileSync('raw_sales_page.html', res.body, 'utf-8');
  console.log("Saved raw_sales_page.html");
}

run().catch(console.error);
