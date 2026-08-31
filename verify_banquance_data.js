import fs from 'fs';

const records = JSON.parse(fs.readFileSync('query_results_north_banquance.json', 'utf-8'));
const northRecords = records.filter(r => r.zone === '北區');

console.log(`Total North records: ${northRecords.length}`);

// Group by Sales Rep
const salesReps = {};
for (const r of northRecords) {
  const rep = r.sales;
  if (!salesReps[rep]) {
    salesReps[rep] = {
      customers: {},
      totalQty: 0,
      totalRtn: 0,
      netQty: 0
    };
  }
  const repObj = salesReps[rep];
  if (!repObj.customers[r.customer]) {
    repObj.customers[r.customer] = {
      products: [],
      custQty: 0,
      custRtn: 0,
      custNet: 0
    };
  }
  const custObj = repObj.customers[r.customer];
  const cleanProduct = r.product.replace(/^\*/, '');
  custObj.products.push({
    rawProduct: r.product,
    cleanProduct: cleanProduct,
    qty: r.qty,
    rtn_qty: r.rtn_qty,
    net: r.qty - r.rtn_qty
  });
  custObj.custQty += r.qty;
  custObj.custRtn += r.rtn_qty;
  custObj.custNet += (r.qty - r.rtn_qty);

  repObj.totalQty += r.qty;
  repObj.totalRtn += r.rtn_qty;
  repObj.netQty += (r.qty - r.rtn_qty);
}

console.log('\n=== 北區 業務與客戶數量及銷貨彙總 ===');
let grandTotalCust = 0;
let grandQty = 0;
let grandRtn = 0;

for (const [rep, data] of Object.entries(salesReps)) {
  const custCount = Object.keys(data.customers).length;
  grandTotalCust += custCount;
  grandQty += data.totalQty;
  grandRtn += data.totalRtn;
  console.log(`\n業務: ${rep} (客戶數: ${custCount})`);
  console.log(`  業務合計 -> 出貨: ${data.totalQty}, 退貨: ${data.totalRtn}, 淨出貨: ${data.netQty}`);
  for (const [cust, cData] of Object.entries(data.customers)) {
    console.log(`    客戶: ${cust} | 筆數: ${cData.products.length} | 出貨: ${cData.custQty} | 退貨: ${cData.custRtn} | 淨出貨: ${cData.custNet}`);
  }
}

console.log('\n=== 全北區總計 ===');
console.log(`業務人數: ${Object.keys(salesReps).length}`);
console.log(`客戶總家數: ${grandTotalCust}`);
console.log(`總出貨: ${grandQty}`);
console.log(`總退貨: ${grandRtn}`);
console.log(`總淨出貨: ${grandQty - grandRtn}`);
