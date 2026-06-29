// Query other review classes to check if they have data
import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const begin = "2025-4-26";
  const end = "2025-6-30";

  const otherClasses = [
    "國中講義:新思維5-6",
    "高中:複習講義"
  ];

  for (const pc of otherClasses) {
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);
    const north = result.records.filter(r => r.zone === '北區');
    console.log(`${pc}: total records = ${result.records.length}, north records = ${north.length}`);
    if (north.length > 0) {
      const qty = north.reduce((s, r) => s + r.qty, 0);
      const rtn = north.reduce((s, r) => s + r.rtn_qty, 0);
      console.log(`  North qty = ${qty}, rtn = ${rtn}, net = ${qty - rtn}`);
    }
  }
}

run().catch(console.error);
