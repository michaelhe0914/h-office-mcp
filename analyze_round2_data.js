import fs from 'fs';

const records = JSON.parse(fs.readFileSync('query_results_round2_review.json', 'utf-8'));

console.log(`Total North records: ${records.length}`);

const salesReps = {};
const productClasses = {};
const overallCustomers = new Set();

for (const r of records) {
  const rep = r.sales || '未指定業務';
  const cust = r.customer;
  const pc = r.productClass;
  const prod = r.product.replace(/^\*/, '');
  const qty = r.qty || 0;
  const rtn = r.rtn_qty || 0;

  overallCustomers.add(cust);

  if (!salesReps[rep]) {
    salesReps[rep] = {
      customers: new Set(),
      qty: 0,
      rtn: 0,
      byClass: {}
    };
  }
  salesReps[rep].customers.add(cust);
  salesReps[rep].qty += qty;
  salesReps[rep].rtn += rtn;

  if (!salesReps[rep].byClass[pc]) {
    salesReps[rep].byClass[pc] = { qty: 0, rtn: 0, products: new Set() };
  }
  salesReps[rep].byClass[pc].qty += qty;
  salesReps[rep].byClass[pc].rtn += rtn;
  salesReps[rep].byClass[pc].products.add(prod);

  if (!productClasses[pc]) {
    productClasses[pc] = { qty: 0, rtn: 0, products: new Set() };
  }
  productClasses[pc].qty += qty;
  productClasses[pc].rtn += rtn;
  productClasses[pc].products.add(prod);
}

console.log(`\nTotal Distinct Customers in North Region: ${overallCustomers.size}`);
console.log(`Total Distinct Sales Reps in North Region: ${Object.keys(salesReps).length}`);

console.log("\n=== Product Classes Summary ===");
for (const [pc, data] of Object.entries(productClasses)) {
  console.log(`Class: ${pc}`);
  console.log(`  Distinct Products: ${data.products.size}`);
  console.log(`  Total Order: ${data.qty}, Return: ${data.rtn}, Net: ${data.qty - data.rtn}`);
}

console.log("\n=== Sales Reps Breakdown ===");
for (const [rep, data] of Object.entries(salesReps)) {
  console.log(`Sales Rep: ${rep}`);
  console.log(`  Associated Customers Count (客戶數量): ${data.customers.size}`);
  console.log(`  Total Order: ${data.qty}, Return: ${data.rtn}, Net: ${data.qty - data.rtn}`);
  for (const [pc, pcData] of Object.entries(data.byClass)) {
    console.log(`    - ${pc}: Order=${pcData.qty}, Return=${pcData.rtn}, Net=${pcData.qty - pcData.rtn}`);
  }
}
