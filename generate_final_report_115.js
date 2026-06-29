import fs from 'fs';

const mainRecords = JSON.parse(fs.readFileSync('query_results_north_115.json', 'utf-8'));

const cleanName = (name) => {
  return name.replace(/^\*/, '').trim();
};

// Merged categorization
const getCategory = (product, productClass) => {
  const cName = cleanName(product);
  if (productClass === "國中講義:新思維5-6") return "新思維系列 (5-6冊)";
  if (productClass === "國中講義:新思維(不含5-6)") return "新思維系列 (不含5-6)";
  if (cName.includes("雙向")) return "雙向複習系列";
  if (cName.includes("735")) return "735輕鬆讀系列";
  if (cName.includes("主題讚")) return "主題讚系列";
  return "其他";
};

// 1. Group by merged name
const mergedMap = {};
for (const r of mainRecords) {
  const cat = getCategory(r.product, r.productClass);
  const pName = cleanName(r.product);
  const key = `${cat} | ${pName}`;
  if (!mergedMap[key]) {
    mergedMap[key] = {
      category: cat,
      product: pName,
      qty: 0,
      rtn: 0
    };
  }
  mergedMap[key].qty += r.qty;
  mergedMap[key].rtn += r.rtn_qty;
}

// 2. Group by raw name
const rawMap = {};
for (const r of mainRecords) {
  const cat = getCategory(r.product, r.productClass);
  const key = `${cat} | ${r.product}`;
  if (!rawMap[key]) {
    rawMap[key] = {
      category: cat,
      product: r.product,
      qty: 0,
      rtn: 0
    };
  }
  rawMap[key].qty += r.qty;
  rawMap[key].rtn += r.rtn_qty;
}

// Output Merged Report
console.log("## scenario_merged");
const categoriesList = ["雙向複習系列", "735輕鬆讀系列", "主題讚系列", "新思維系列 (不含5-6)", "新思維系列 (5-6冊)"];

let grandQty = 0;
let grandRtn = 0;

for (const cat of categoriesList) {
  console.log(`\n### ${cat}`);
  console.log(`| 產品名稱 | 訂量 | 退量 | 淨出貨 | 退書率 |`);
  console.log(`| :--- | ---: | ---: | ---: | ---: |`);
  
  const items = Object.values(mergedMap).filter(x => x.category === cat).sort((a,b)=>a.product.localeCompare(b.product));
  let catQty = 0;
  let catRtn = 0;
  for (const item of items) {
    const net = item.qty - item.rtn;
    const rate = item.qty > 0 ? ((item.rtn / (item.qty + item.rtn)) * 100).toFixed(1) + '%' : '0.0%';
    console.log(`| ${item.product} | ${item.qty.toLocaleString()} | ${item.rtn.toLocaleString()} | **${net.toLocaleString()}** | ${rate} |`);
    catQty += item.qty;
    catRtn += item.rtn;
  }
  const catNet = catQty - catRtn;
  const catRate = catQty > 0 ? ((catRtn / (catQty + catRtn)) * 100).toFixed(1) + '%' : '0.0%';
  console.log(`| **${cat} 小計** | **${catQty.toLocaleString()}** | **${catRtn.toLocaleString()}** | **${catNet.toLocaleString()}** | **${catRate}** |`);
  
  if (cat !== "新思維系列 (5-6冊)") {
    grandQty += catQty;
    grandRtn += catRtn;
  }
}
const grandNet = grandQty - grandRtn;
const grandRate = grandQty > 0 ? ((grandRtn / (grandQty + grandRtn)) * 100).toFixed(1) + '%' : '0.0%';
console.log(`\n| **總計 (不含5-6)** | **${grandQty.toLocaleString()}** | **${grandRtn.toLocaleString()}** | **${grandNet.toLocaleString()}** | **${grandRate}** |`);


// Output Raw Report
console.log("\n## scenario_raw");
for (const cat of categoriesList) {
  console.log(`\n### ${cat} (原始細目)`);
  console.log(`| 產品名稱 (原名稱) | 訂量 | 退量 | 淨出貨 | 退書率 |`);
  console.log(`| :--- | ---: | ---: | ---: | ---: |`);
  
  const items = Object.values(rawMap).filter(x => x.category === cat).sort((a,b)=>a.product.localeCompare(b.product));
  let catQty = 0;
  let catRtn = 0;
  for (const item of items) {
    const net = item.qty - item.rtn;
    const rate = item.qty > 0 ? ((item.rtn / (item.qty + item.rtn)) * 100).toFixed(1) + '%' : '0.0%';
    console.log(`| ${item.product} | ${item.qty.toLocaleString()} | ${item.rtn.toLocaleString()} | **${net.toLocaleString()}** | ${rate} |`);
    catQty += item.qty;
    catRtn += item.rtn;
  }
  const catNet = catQty - catRtn;
  const catRate = catQty > 0 ? ((catRtn / (catQty + catRtn)) * 100).toFixed(1) + '%' : '0.0%';
  console.log(`| **${cat} 原始小計** | **${catQty.toLocaleString()}** | **${catRtn.toLocaleString()}** | **${catNet.toLocaleString()}** | **${catRate}** |`);
}
