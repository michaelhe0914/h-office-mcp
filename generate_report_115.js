import fs from 'fs';

const mainRecords = JSON.parse(fs.readFileSync('query_results_north_115.json', 'utf-8'));

const cleanName = (name) => {
  return name.replace(/^\*/, '').trim();
};

const summaryRaw = {};
const summaryMerged = {};

for (const r of mainRecords) {
  const rawKey = `${r.productClass} | ${r.product}`;
  if (!summaryRaw[rawKey]) {
    summaryRaw[rawKey] = {
      productClass: r.productClass,
      product: r.product,
      qty: 0,
      rtn: 0
    };
  }
  summaryRaw[rawKey].qty += r.qty;
  summaryRaw[rawKey].rtn += r.rtn_qty;

  const mergedKey = `${r.productClass} | ${cleanName(r.product)}`;
  if (!summaryMerged[mergedKey]) {
    summaryMerged[mergedKey] = {
      productClass: r.productClass,
      product: cleanName(r.product),
      qty: 0,
      rtn: 0
    };
  }
  summaryMerged[mergedKey].qty += r.qty;
  summaryMerged[mergedKey].rtn += r.rtn_qty;
}

console.log("=== MERGED PRODUCTS ===");
for (const [key, val] of Object.entries(summaryMerged).sort((a,b)=>a[0].localeCompare(b[0]))) {
  console.log(`${val.productClass} | ${val.product}: 訂=${val.qty}, 退=${val.rtn}, 淨=${val.qty - val.rtn}`);
}
