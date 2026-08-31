import fs from 'fs';

const records = JSON.parse(fs.readFileSync('query_results_round2_review.json', 'utf-8'));

const repCustomers = {};

for (const r of records) {
  const rep = r.sales || '未指定業務';
  if (!repCustomers[rep]) repCustomers[rep] = new Set();
  repCustomers[rep].add(r.customer);
}

for (const [rep, custs] of Object.entries(repCustomers)) {
  const sorted = [...custs].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  console.log(`\n業務: ${rep} (${sorted.length} 家客戶)`);
  console.log(`客戶列表: ${sorted.join('、')}`);
}
