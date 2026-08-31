import fs from 'fs';

const data = JSON.parse(fs.readFileSync('query_results_north_banquance.json', 'utf-8'));
const north = data.filter(r => r.zone === '北區');

console.log('Total North records:', north.length);

const reps = {};
for (const r of north) {
  if (!reps[r.sales]) {
    reps[r.sales] = {
      customers: new Set(),
      totalQty: 0,
      totalRtn: 0,
      products: new Set(),
      customerDetails: {}
    };
  }
  const repObj = reps[r.sales];
  repObj.customers.add(r.customer);
  repObj.totalQty += r.qty;
  repObj.totalRtn += r.rtn_qty;
  repObj.products.add(r.product);

  if (!repObj.customerDetails[r.customer]) {
    repObj.customerDetails[r.customer] = { qty: 0, rtn: 0, items: [] };
  }
  repObj.customerDetails[r.customer].qty += r.qty;
  repObj.customerDetails[r.customer].rtn += r.rtn_qty;
  repObj.customerDetails[r.customer].items.push(r);
}

console.log('\n--- 北區各業務與所屬客戶數量 ---');
for (const [rep, info] of Object.entries(reps)) {
  console.log(`業務: ${rep} | 客戶數量: ${info.customers.size} | 出貨量: ${info.totalQty} | 退貨量: ${info.totalRtn} | 淨出貨: ${info.totalQty - info.totalRtn}`);
  console.log('  客戶列表:', [...info.customers].join(', '));
}

console.log('\nAll 北區 Products:', [...new Set(north.map(r => r.product))]);
