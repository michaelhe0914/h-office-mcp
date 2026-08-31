import fs from 'fs';

const data = JSON.parse(fs.readFileSync('full_aggregated_data.json', 'utf-8'));

function formatTable(items, seriesName) {
  let md = `\n### 📘 ${seriesName} (共 ${items.length} 項)\n\n`;
  md += `| 品項名稱 | 訂量 | 退量 | 淨出貨 |\n`;
  md += `| :--- | ---: | ---: | ---: |\n`;
  
  // Sort items: first by subject/volume, or natural name
  const sorted = [...items].sort((a, b) => a.product.localeCompare(b.product, 'zh-TW'));
  let totalQty = 0;
  let totalRtn = 0;

  for (const item of sorted) {
    totalQty += item.qty;
    totalRtn += item.rtn_qty;
    const net = item.qty - item.rtn_qty;
    md += `| ${item.product} | ${item.qty.toLocaleString()} | ${item.rtn_qty.toLocaleString()} | **${net.toLocaleString()}** |\n`;
  }
  const netTotal = totalQty - totalRtn;
  md += `| **${seriesName} 小計** | **${totalQty.toLocaleString()}** | **${totalRtn.toLocaleString()}** | **${netTotal.toLocaleString()}** |\n`;
  return md;
}

let fullDoc = `# 📊 北區單冊講義、單冊卷各品項出貨統計表 (3/26 ~ 8/28)\n\n`;
fullDoc += `**查詢區間**：2026/03/26 ～ 2026/08/28 (民國115學年度)\n`;
fullDoc += `**統計範圍**：北區業務團隊（單冊講義、單冊卷）\n\n`;

fullDoc += `## 📌 一、 總計彙整表\n\n`;
fullDoc += `| 大類別 | 系列名稱 | 品項數 | 訂量 (本) | 退量 (本) | 實際出貨 (本) |\n`;
fullDoc += `| :--- | :--- | ---: | ---: | ---: | ---: |\n`;

const s_sx = data.jiangyi.shuangXiang;
const s_735 = data.jiangyi.sevenThreeFive;
const s_st = data.jiangyi.shiTi;
const s_xj = data.jiangyi.xinJiangYi;

const s_kang = data.exam.kang;
const s_nan = data.exam.nan;
const s_han = data.exam.han;
const s_white = data.exam.white;
const s_12k = data.exam.exam12K;

function sumList(list) {
  const qty = list.reduce((s, x) => s + x.qty, 0);
  const rtn = list.reduce((s, x) => s + x.rtn_qty, 0);
  return { qty, rtn, net: qty - rtn, count: list.length };
}

const c_sx = sumList(s_sx);
const c_735 = sumList(s_735);
const c_st = sumList(s_st);
const c_xj = sumList(s_xj);
const totalJY = { qty: c_sx.qty + c_735.qty + c_st.qty + c_xj.qty, rtn: c_sx.rtn + c_735.rtn + c_st.rtn + c_xj.rtn, count: c_sx.count + c_735.count + c_st.count + c_xj.count };

const c_kang = sumList(s_kang);
const c_nan = sumList(s_nan);
const c_han = sumList(s_han);
const c_white = sumList(s_white);
const c_12k = sumList(s_12k);
const totalEX = { qty: c_kang.qty + c_nan.qty + c_han.qty + c_white.qty + c_12k.qty, rtn: c_kang.rtn + c_nan.rtn + c_han.rtn + c_white.rtn + c_12k.rtn, count: c_kang.count + c_nan.count + c_han.count + c_white.count + c_12k.count };

fullDoc += `| **國中單冊講義** | 雙向講義系列 | ${c_sx.count} | ${c_sx.qty.toLocaleString()} | ${c_sx.rtn.toLocaleString()} | **${c_sx.net.toLocaleString()}** |\n`;
fullDoc += `| | 735講義系列 | ${c_735.count} | ${c_735.qty.toLocaleString()} | ${c_735.rtn.toLocaleString()} | **${c_735.net.toLocaleString()}** |\n`;
fullDoc += `| | 試題篇系列 | ${c_st.count} | ${c_st.qty.toLocaleString()} | ${c_st.rtn.toLocaleString()} | **${c_st.net.toLocaleString()}** |\n`;
fullDoc += `| | 新講義系列 | ${c_xj.count} | ${c_xj.qty.toLocaleString()} | ${c_xj.rtn.toLocaleString()} | **${c_xj.net.toLocaleString()}** |\n`;
fullDoc += `| | **單冊講義 小計** | **${totalJY.count}** | **${totalJY.qty.toLocaleString()}** | **${totalJY.rtn.toLocaleString()}** | **${(totalJY.qty - totalJY.rtn).toLocaleString()}** |\n`;
fullDoc += `| **國中單冊考卷** | 康卷系列 (A/B卷) | ${c_kang.count} | ${c_kang.qty.toLocaleString()} | ${c_kang.rtn.toLocaleString()} | **${c_kang.net.toLocaleString()}** |\n`;
fullDoc += `| | 南卷系列 | ${c_nan.count} | ${c_nan.qty.toLocaleString()} | ${c_nan.rtn.toLocaleString()} | **${c_nan.net.toLocaleString()}** |\n`;
fullDoc += `| | 翰卷系列 (A/B卷) | ${c_han.count} | ${c_han.qty.toLocaleString()} | ${c_han.rtn.toLocaleString()} | **${c_han.net.toLocaleString()}** |\n`;
fullDoc += `| | 8K單冊白卷系列 | ${c_white.count} | ${c_white.qty.toLocaleString()} | ${c_white.rtn.toLocaleString()} | **${c_white.net.toLocaleString()}** |\n`;
fullDoc += `| | 12K/9K單冊考卷 | ${c_12k.count} | ${c_12k.qty.toLocaleString()} | ${c_12k.rtn.toLocaleString()} | **${c_12k.net.toLocaleString()}** |\n`;
fullDoc += `| | **單冊考卷 小計** | **${totalEX.count}** | **${totalEX.qty.toLocaleString()}** | **${totalEX.rtn.toLocaleString()}** | **${(totalEX.qty - totalEX.rtn).toLocaleString()}** |\n`;
fullDoc += `| **總計** | **全部合計** | **${totalJY.count + totalEX.count}** | **${(totalJY.qty + totalEX.qty).toLocaleString()}** | **${(totalJY.rtn + totalEX.rtn).toLocaleString()}** | **${((totalJY.qty - totalJY.rtn) + (totalEX.qty - totalEX.rtn)).toLocaleString()}** |\n\n`;

fullDoc += `---\n\n## 📖 二、 國中單冊講義 各品項明細清單\n`;
fullDoc += formatTable(s_sx, '雙向講義系列');
fullDoc += formatTable(s_735, '735講義系列');
fullDoc += formatTable(s_st, '試題篇講義系列');
fullDoc += formatTable(s_xj, '新講義系列');

fullDoc += `---\n\n## 📝 三、 國中單冊考卷 各品項明細清單\n`;
fullDoc += formatTable(s_kang, '康卷系列 (A卷 / B卷)');
fullDoc += formatTable(s_nan, '南卷系列');
fullDoc += formatTable(s_han, '翰卷系列 (A卷 / B卷)');
fullDoc += formatTable(s_white, '8K單冊白卷系列');
fullDoc += formatTable(s_12k, '12K/9K單冊考卷系列');

fs.writeFileSync('north_single_volume_summary_326_828.md', fullDoc, 'utf-8');
console.log('Successfully generated north_single_volume_summary_326_828.md!');
