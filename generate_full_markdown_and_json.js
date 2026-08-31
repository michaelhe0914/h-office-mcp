import fs from 'fs';

const jiangyi = JSON.parse(fs.readFileSync('temp_jiangyi_2026.json', 'utf-8'));
const exam = JSON.parse(fs.readFileSync('temp_exam_2026.json', 'utf-8'));

// Helper to aggregate by product
function aggregateByProduct(records) {
  const map = {};
  for (const r of records) {
    if (!map[r.product]) {
      map[r.product] = {
        product: r.product,
        productClass: r.productClass,
        qty: 0,
        rtn_qty: 0,
        net_qty: 0,
      };
    }
    map[r.product].qty += r.qty;
    map[r.product].rtn_qty += r.rtn_qty;
    map[r.product].net_qty = map[r.product].qty - map[r.product].rtn_qty;
  }
  return Object.values(map);
}

// Check all jiangyi series
const jiangyiShuangXiang = jiangyi.filter(r => r.productClass.includes('雙向'));
const jiangyi735 = jiangyi.filter(r => r.productClass.includes('735'));
const jiangyiShiTi = jiangyi.filter(r => r.productClass.includes('試題'));
const jiangyiXin = jiangyi.filter(r => r.productClass.includes('新講義'));

// Check all exam series
const examKang = exam.filter(r => r.productClass.includes('康卷'));
const examNan = exam.filter(r => r.productClass.includes('南卷'));
const examHan = exam.filter(r => r.productClass.includes('翰卷'));
const examWhite = exam.filter(r => r.productClass.includes('白卷'));
const exam12K = exam.filter(r => r.productClass.includes('12K'));

const summary = {
  dateRange: '2026-03-26 ~ 2026-08-28 (115學年)',
  region: '北區',
  jiangyi: {
    shuangXiang: aggregateByProduct(jiangyiShuangXiang),
    sevenThreeFive: aggregateByProduct(jiangyi735),
    shiTi: aggregateByProduct(jiangyiShiTi),
    xinJiangYi: aggregateByProduct(jiangyiXin),
  },
  exam: {
    kang: aggregateByProduct(examKang),
    nan: aggregateByProduct(examNan),
    han: aggregateByProduct(examHan),
    white: aggregateByProduct(examWhite),
    exam12K: aggregateByProduct(exam12K),
  }
};

fs.writeFileSync('full_aggregated_data.json', JSON.stringify(summary, null, 2), 'utf-8');

console.log('=== SUMMARY OF TOTALS ===');
function printCatSummary(name, list) {
  const qty = list.reduce((s, x) => s + x.qty, 0);
  const rtn = list.reduce((s, x) => s + x.rtn_qty, 0);
  const net = qty - rtn;
  const count = list.length;
  console.log(`${name}: 品項數=${count}, 訂量=${qty}, 退量=${rtn}, 淨出貨=${net}`);
  return { qty, rtn, net, count };
}

console.log('--- 單冊講義 ---');
const s_sx = printCatSummary('雙向講義系列', summary.jiangyi.shuangXiang);
const s_735 = printCatSummary('735講義系列', summary.jiangyi.sevenThreeFive);
const s_st = printCatSummary('試題篇系列', summary.jiangyi.shiTi);
const s_xj = printCatSummary('新講義系列', summary.jiangyi.xinJiangYi);
const totalJiangyiQty = s_sx.qty + s_735.qty + s_st.qty + s_xj.qty;
const totalJiangyiRtn = s_sx.rtn + s_735.rtn + s_st.rtn + s_xj.rtn;
console.log(`>>> 單冊講義 合計: 訂量=${totalJiangyiQty}, 退量=${totalJiangyiRtn}, 淨出貨=${totalJiangyiQty - totalJiangyiRtn}, 總品項數=${s_sx.count + s_735.count + s_st.count + s_xj.count}`);

console.log('\n--- 單冊卷 ---');
const s_kang = printCatSummary('康卷系列(A/B卷)', summary.exam.kang);
const s_nan = printCatSummary('南卷系列', summary.exam.nan);
const s_han = printCatSummary('翰卷系列(A/B卷)', summary.exam.han);
const s_white = printCatSummary('8K單冊白卷系列', summary.exam.white);
const s_12k = printCatSummary('12K/9K單冊考卷系列', summary.exam.exam12K);
const totalExamQty = s_kang.qty + s_nan.qty + s_han.qty + s_white.qty + s_12k.qty;
const totalExamRtn = s_kang.rtn + s_nan.rtn + s_han.rtn + s_white.rtn + s_12k.rtn;
console.log(`>>> 單冊卷 合計: 訂量=${totalExamQty}, 退量=${totalExamRtn}, 淨出貨=${totalExamQty - totalExamRtn}, 總品項數=${s_kang.count + s_nan.count + s_han.count + s_white.count + s_12k.count}`);

console.log(`\n========================================`);
console.log(`>>> 北區 3/26~8/28 總訂量: ${totalJiangyiQty + totalExamQty}`);
console.log(`>>> 北區 3/26~8/28 總退量: ${totalJiangyiRtn + totalExamRtn}`);
console.log(`>>> 北區 3/26~8/28 總淨出貨: ${(totalJiangyiQty - totalJiangyiRtn) + (totalExamQty - totalExamRtn)}`);
console.log(`========================================`);
