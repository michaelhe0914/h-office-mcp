import fs from 'fs';

const records = JSON.parse(fs.readFileSync('query_results_north_115.json', 'utf-8'));

const summary = {};
for (const r of records) {
  if (!summary[r.productClass]) {
    summary[r.productClass] = { qty: 0, rtn: 0 };
  }
  summary[r.productClass].qty += r.qty;
  summary[r.productClass].rtn += r.rtn_qty;
}

console.log("=== SUMMARY BY PRODUCT CLASS (115) ===");
for (const [pc, data] of Object.entries(summary)) {
  console.log(`${pc}: 訂量=${data.qty}, 退量=${data.rtn}, 淨出貨=${data.qty - data.rtn}`);
}
