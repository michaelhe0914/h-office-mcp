import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const begin = "2026-5-5";
  const end = "2026-6-5";
  const pc = "國中講義:新思維(不含5-6)";

  const params = { begin, end, product_class: pc };
  const res = await request(BASE_URL, "/sales", { params });
  const result = parseSalesPage(res.body);

  const northRecords = result.records.filter(r => r.zone === "北區");
  let qty = 0;
  let rtn = 0;
  for (const r of northRecords) {
    qty += r.qty;
    rtn += r.rtn_qty;
  }
  console.log(`Class: ${pc} -> qty: ${qty}, rtn: ${rtn}`);
}

run().catch(console.error);
