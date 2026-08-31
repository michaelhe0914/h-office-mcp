import fs from 'fs';

const records = JSON.parse(fs.readFileSync('query_results_hs_tw.json', 'utf-8'));

const repInfo = new Set();
const products = new Set();
const customers = new Set();

for (const r of records) {
  repInfo.add(`${r.zone} | ${r.sales}`);
  products.add(`${r.productClass} | ${r.product}`);
  customers.add(r.customer);
}

console.log("=== Sales Reps in Data ===");
[...repInfo].sort().forEach(r => console.log(r));

console.log("\n=== Products in Data ===");
[...products].sort().forEach(p => console.log(p));

console.log(`\n=== Total Customers: ${customers.size} ===`);
console.log("Customers:");
[...customers].sort().forEach(c => console.log(`- ${c}`));
