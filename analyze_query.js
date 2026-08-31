import fs from 'fs';

const jiangyi = JSON.parse(fs.readFileSync('temp_jiangyi_2026.json', 'utf-8'));
const exam = JSON.parse(fs.readFileSync('temp_exam_2026.json', 'utf-8'));

console.log('=== 業務人員名單 ===');
const reps = new Set();
jiangyi.forEach(r => reps.add(r.sales));
exam.forEach(r => reps.add(r.sales));
console.log([...reps].join(', '));

console.log('\n=== 單冊講義：各類別加總 ===');
const jiangyiByClass = {};
jiangyi.forEach(r => {
  if (!jiangyiByClass[r.productClass]) jiangyiByClass[r.productClass] = { qty: 0, rtn: 0, count: 0 };
  jiangyiByClass[r.productClass].qty += r.qty;
  jiangyiByClass[r.productClass].rtn += r.rtn_qty;
  jiangyiByClass[r.productClass].count++;
});
for (const [k, v] of Object.entries(jiangyiByClass)) {
  console.log(`${k}: 訂量=${v.qty}, 退量=${v.rtn}, 淨量=${v.qty - v.rtn} (筆數: ${v.count})`);
}

console.log('\n=== 單冊卷：各類別加總 ===');
const examByClass = {};
exam.forEach(r => {
  if (!examByClass[r.productClass]) examByClass[r.productClass] = { qty: 0, rtn: 0, count: 0 };
  examByClass[r.productClass].qty += r.qty;
  examByClass[r.productClass].rtn += r.rtn_qty;
  examByClass[r.productClass].count++;
});
for (const [k, v] of Object.entries(examByClass)) {
  console.log(`${k}: 訂量=${v.qty}, 退量=${v.rtn}, 淨量=${v.qty - v.rtn} (筆數: ${v.count})`);
}

// Group by series
console.log('\n=============================================');
console.log('=== 單冊講義 各品項明細彙整 ===');
console.log('=============================================');

function analyzeProducts(records, title) {
  console.log(`\n### ${title}`);
  const prodMap = {};
  for (const r of records) {
    if (!prodMap[r.product]) prodMap[r.product] = { class: r.productClass, qty: 0, rtn: 0 };
    prodMap[r.product].qty += r.qty;
    prodMap[r.product].rtn += r.rtn_qty;
  }
  let totalQty = 0;
  let totalRtn = 0;
  const sorted = Object.entries(prodMap).sort((a, b) => a[0].localeCompare(b[0], 'zh-TW'));
  for (const [p, d] of sorted) {
    const net = d.qty - d.rtn;
    console.log(`- ${p} | 訂量: ${d.qty} | 退量: ${d.rtn} | 淨出貨: ${net}`);
    totalQty += d.qty;
    totalRtn += d.rtn;
  }
  console.log(`小計 -> 訂量: ${totalQty} | 退量: ${totalRtn} | 淨出貨: ${totalQty - totalRtn} | 品項數: ${sorted.length}`);
  return { totalQty, totalRtn, net: totalQty - totalRtn, count: sorted.length };
}

analyzeProducts(jiangyi.filter(r => r.productClass.includes('雙向')), '雙向講義系列');
analyzeProducts(jiangyi.filter(r => r.productClass.includes('735')), '735講義系列');
analyzeProducts(jiangyi.filter(r => r.productClass.includes('試題')), '試題篇講義系列');
analyzeProducts(jiangyi.filter(r => r.productClass.includes('新講義')), '新講義系列');

console.log('\n=============================================');
console.log('=== 單冊卷 各品項明細彙整 ===');
console.log('=============================================');
analyzeProducts(exam.filter(r => r.productClass.includes('康卷')), '康卷系列');
analyzeProducts(exam.filter(r => r.productClass.includes('南卷')), '南卷系列');
analyzeProducts(exam.filter(r => r.productClass.includes('翰卷')), '翰卷系列');
analyzeProducts(exam.filter(r => r.productClass.includes('白卷')), '8K單冊白卷系列');
analyzeProducts(exam.filter(r => r.productClass.includes('12K')), '12K考卷系列');
