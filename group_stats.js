import fs from 'fs';

const jiangyi = JSON.parse(fs.readFileSync('temp_jiangyi_2026.json', 'utf-8'));
const exam = JSON.parse(fs.readFileSync('temp_exam_2026.json', 'utf-8'));

console.log('=== 講義依版本分類 ===');
// Version classification for Handouts
// 康版: productClass or product has 康版
// 南版: productClass or product has 南版
// 翰版: productClass or product has 翰版
// 綜合版 / 新講義: productClass has 新講義 or product has 新講義
const jiangyiByVersion = {
  '康版 (K版)': { qty: 0, rtn: 0, count: 0, series: {} },
  '南版 (N版)': { qty: 0, rtn: 0, count: 0, series: {} },
  '翰版 (H版)': { qty: 0, rtn: 0, count: 0, series: {} },
  '綜合版 (新講義)': { qty: 0, rtn: 0, count: 0, series: {} }
};

for (const r of jiangyi) {
  let v = '綜合版 (新講義)';
  if (r.productClass.includes('康版') || r.product.includes('康版')) v = '康版 (K版)';
  else if (r.productClass.includes('南版') || r.product.includes('南版')) v = '南版 (N版)';
  else if (r.productClass.includes('翰版') || r.product.includes('翰版')) v = '翰版 (H版)';
  
  let s = '新講義';
  if (r.productClass.includes('雙向')) s = '雙向';
  else if (r.productClass.includes('735')) s = '735';
  else if (r.productClass.includes('試題')) s = '試題篇';

  jiangyiByVersion[v].qty += r.qty;
  jiangyiByVersion[v].rtn += r.rtn_qty;
  jiangyiByVersion[v].count++;
  if (!jiangyiByVersion[v].series[s]) jiangyiByVersion[v].series[s] = { qty: 0, rtn: 0 };
  jiangyiByVersion[v].series[s].qty += r.qty;
  jiangyiByVersion[v].series[s].rtn += r.rtn_qty;
}

console.log(JSON.stringify(jiangyiByVersion, null, 2));

console.log('\n=== 考卷依種類分類 ===');
// Exam kinds:
// 1. By Type (A卷, B卷, 南卷, 白卷, 12K/9K卷)
// 2. By Product Class (康卷-A卷, 康卷-B卷, 南卷, 翰卷-A卷, 翰卷-B卷, 8k單冊白卷, 12K考卷)
const examByType = {
  'A卷 (含康A、翰A)': { qty: 0, rtn: 0, count: 0 },
  'B卷 (含康B、翰B)': { qty: 0, rtn: 0, count: 0 },
  '南卷 (標準卷)': { qty: 0, rtn: 0, count: 0 },
  '8K單冊白卷': { qty: 0, rtn: 0, count: 0 },
  '12K/9K考卷': { qty: 0, rtn: 0, count: 0 }
};

const examByDetailedType = {
  '康卷 - A卷': { qty: 0, rtn: 0, count: 0 },
  '康卷 - B卷': { qty: 0, rtn: 0, count: 0 },
  '南卷 (標準卷)': { qty: 0, rtn: 0, count: 0 },
  '翰卷 - A卷': { qty: 0, rtn: 0, count: 0 },
  '翰卷 - B卷': { qty: 0, rtn: 0, count: 0 },
  '8K單冊白卷': { qty: 0, rtn: 0, count: 0 },
  '12K/9K考卷': { qty: 0, rtn: 0, count: 0 }
};

for (const r of exam) {
  if (r.productClass.includes('康卷-A卷')) {
    examByType['A卷 (含康A、翰A)'].qty += r.qty;
    examByType['A卷 (含康A、翰A)'].rtn += r.rtn_qty;
    examByType['A卷 (含康A、翰A)'].count++;

    examByDetailedType['康卷 - A卷'].qty += r.qty;
    examByDetailedType['康卷 - A卷'].rtn += r.rtn_qty;
    examByDetailedType['康卷 - A卷'].count++;
  } else if (r.productClass.includes('康卷-B卷')) {
    examByType['B卷 (含康B、翰B)'].qty += r.qty;
    examByType['B卷 (含康B、翰B)'].rtn += r.rtn_qty;
    examByType['B卷 (含康B、翰B)'].count++;

    examByDetailedType['康卷 - B卷'].qty += r.qty;
    examByDetailedType['康卷 - B卷'].rtn += r.rtn_qty;
    examByDetailedType['康卷 - B卷'].count++;
  } else if (r.productClass.includes('翰卷-A卷')) {
    examByType['A卷 (含康A、翰A)'].qty += r.qty;
    examByType['A卷 (含康A、翰A)'].rtn += r.rtn_qty;
    examByType['A卷 (含康A、翰A)'].count++;

    examByDetailedType['翰卷 - A卷'].qty += r.qty;
    examByDetailedType['翰卷 - A卷'].rtn += r.rtn_qty;
    examByDetailedType['翰卷 - A卷'].count++;
  } else if (r.productClass.includes('翰卷-B卷')) {
    examByType['B卷 (含康B、翰B)'].qty += r.qty;
    examByType['B卷 (含康B、翰B)'].rtn += r.rtn_qty;
    examByType['B卷 (含康B、翰B)'].count++;

    examByDetailedType['翰卷 - B卷'].qty += r.qty;
    examByDetailedType['翰卷 - B卷'].rtn += r.rtn_qty;
    examByDetailedType['翰卷 - B卷'].count++;
  } else if (r.productClass.includes('南卷')) {
    examByType['南卷 (標準卷)'].qty += r.qty;
    examByType['南卷 (標準卷)'].rtn += r.rtn_qty;
    examByType['南卷 (標準卷)'].count++;

    examByDetailedType['南卷 (標準卷)'].qty += r.qty;
    examByDetailedType['南卷 (標準卷)'].rtn += r.rtn_qty;
    examByDetailedType['南卷 (標準卷)'].count++;
  } else if (r.productClass.includes('白卷')) {
    examByType['8K單冊白卷'].qty += r.qty;
    examByType['8K單冊白卷'].rtn += r.rtn_qty;
    examByType['8K單冊白卷'].count++;

    examByDetailedType['8K單冊白卷'].qty += r.qty;
    examByDetailedType['8K單冊白卷'].rtn += r.rtn_qty;
    examByDetailedType['8K單冊白卷'].count++;
  } else if (r.productClass.includes('12K')) {
    examByType['12K/9K考卷'].qty += r.qty;
    examByType['12K/9K考卷'].rtn += r.rtn_qty;
    examByType['12K/9K考卷'].count++;

    examByDetailedType['12K/9K考卷'].qty += r.qty;
    examByDetailedType['12K/9K考卷'].rtn += r.rtn_qty;
    examByDetailedType['12K/9K考卷'].count++;
  }
}

console.log('Exam by general type:', JSON.stringify(examByType, null, 2));
console.log('Exam by detailed type:', JSON.stringify(examByDetailedType, null, 2));
