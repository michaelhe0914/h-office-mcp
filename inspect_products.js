import fs from 'fs';

const records = JSON.parse(fs.readFileSync('query_results_hs_tw.json', 'utf-8'));

const productsByClass = {};

for (const r of records) {
  if (!productsByClass[r.productClass]) {
    productsByClass[r.productClass] = {};
  }
  if (!productsByClass[r.productClass][r.product]) {
    productsByClass[r.productClass][r.product] = { qty: 0, rtn: 0 };
  }
  productsByClass[r.productClass][r.product].qty += r.qty;
  productsByClass[r.productClass][r.product].rtn += r.rtn_qty;
}

console.log("=== Products by Class ===");
for (const [pc, prods] of Object.entries(productsByClass)) {
  console.log(`\nClass: ${pc}`);
  Object.entries(prods).forEach(([name, count]) => {
    console.log(`  - ${name} | Qty: ${count.qty} | Rtn: ${count.rtn}`);
  });
}
