// Verify data integrity: check total quantities per sales rep against initial query summary
import fs from 'fs';

const allRecords = JSON.parse(fs.readFileSync('query_results_north_single_volume.json', 'utf-8'));

const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉偉", "朱鵬學"];

for (const sales of targetSales) {
  const recs = allRecords.filter(r => r.sales === sales);
  const totalQty = recs.reduce((sum, r) => sum + r.qty, 0);
  console.log(`${sales}: ${recs.length} records, total qty: ${totalQty}`);
  
  // Check by category
  const categories = ['雙向', '735', '試題篇', '新講義'];
  for (const cat of categories) {
    const catFilter = (r) => {
      if (cat === '雙向') return r.productClass.includes('雙向');
      if (cat === '735') return r.productClass.includes('735');
      if (cat === '試題篇') return r.productClass.includes('試題');
      if (cat === '新講義') return r.productClass.includes('新講義');
      return false;
    };
    const catRecs = recs.filter(catFilter);
    const catQty = catRecs.reduce((sum, r) => sum + r.qty, 0);
    if (catQty > 0) console.log(`  ${cat}: qty ${catQty}`);
  }
}

// Also verify specific customer "柏興圖書" from 何光傑
console.log("\n=== Verification: 柏興圖書 (何光傑) ===");
const pbRecs = allRecords.filter(r => r.sales === "何光傑" && r.customer === "柏興圖書");
console.log(`Total records: ${pbRecs.length}, Total qty: ${pbRecs.reduce((s, r) => s + r.qty, 0)}`);

// Sum by category
const catSums = {};
for (const r of pbRecs) {
  let cat;
  if (r.productClass.includes('雙向')) cat = '雙向';
  else if (r.productClass.includes('735')) cat = '735';
  else if (r.productClass.includes('試題')) cat = '試題篇';
  else if (r.productClass.includes('新講義')) cat = '新講義';
  if (!catSums[cat]) catSums[cat] = 0;
  catSums[cat] += r.qty;
}
console.log("By category:", catSums);

// Check: 雙向 total from xlsx should be 2341
// Sum: N=115, K=1091, H=1135 = 2341 ✓
const sxRecs = pbRecs.filter(r => r.productClass.includes('雙向'));
console.log(`雙向 total: ${sxRecs.reduce((s, r) => s + r.qty, 0)}`);
