import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const begin = "2026-4-26";
  const end = "2026-6-30";

  const potentialClasses = [
    "國中講義:基精123",
    "國中講義:輔導1-2",
    "國中講義:考前30天",
    "國中講義:3900題",
    "國中講義:歷屆試題"
  ];

  for (const pc of potentialClasses) {
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);
    const north = result.records.filter(r => r.zone === '北區');
    console.log(`${pc}: total records = ${result.records.length}, north records = ${north.length}`);
    if (north.length > 0) {
      const qty = north.reduce((s, r) => s + r.qty, 0);
      const rtn = north.reduce((s, r) => s + r.rtn_qty, 0);
      console.log(`  North qty = ${qty}, rtn = ${rtn}, net = ${qty - rtn}`);
      // Show products
      const counts = {};
      for (const r of north) {
        if (!counts[r.product]) counts[r.product] = 0;
        counts[r.product] += (r.qty - r.rtn_qty);
      }
      console.log("  Products net:", counts);
    }
  }
}

run().catch(console.error);
