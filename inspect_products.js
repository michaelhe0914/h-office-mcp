import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const begin = "2025-4-26";
  const end = "2025-6-30";
  const pc = "國中講義:新思維5-6";
  
  const params = { begin, end, product_class: pc };
  const res = await request(BASE_URL, "/sales", { params });
  const result = parseSalesPage(res.body);
  const north = result.records.filter(r => r.zone === '北區');
  
  console.log("=== Products in 新思維5-6 ===");
  const counts = {};
  for (const r of north) {
    if (!counts[r.product]) counts[r.product] = { qty: 0, rtn: 0 };
    counts[r.product].qty += r.qty;
    counts[r.product].rtn += r.rtn_qty;
  }
  for (const [p, data] of Object.entries(counts)) {
    console.log(`- ${p}: 訂=${data.qty}, 退=${data.rtn}, 淨=${data.qty - data.rtn}`);
  }
}

run().catch(console.error);
