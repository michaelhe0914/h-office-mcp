import fs from 'fs';

const records = JSON.parse(fs.readFileSync('query_results_hs_tw.json', 'utf-8'));

console.log("=== Customer Sales Summary ===");
const summary = {};
for (const r of records) {
  if (!summary[r.customer]) {
    summary[r.customer] = { zone: r.zone, sales: r.sales, qty: 0, rtn_qty: 0 };
  }
  summary[r.customer].qty += r.qty;
  summary[r.customer].rtn_qty += r.rtn_qty;
}

Object.entries(summary).sort((a, b) => b[1].qty - a[1].qty).forEach(([c, s]) => {
  console.log(`${c} | Zone: ${s.zone} | Sales: ${s.sales} | Qty: ${s.qty} | Rtn: ${s.rtn_qty}`);
});
