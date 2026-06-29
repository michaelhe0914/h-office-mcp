import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const begin = "2026-6-8";
  const end = "2026-6-9";

  // 1. Fetch sales page to get all product classes
  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  
  // Find all classes containing "新思維"
  const targetClasses = parsedInitial.productClasses
    .map(c => c.name)
    .filter(name => name.includes("新思維"));

  console.log(`Found target classes:`, targetClasses);

  let grandTotalQty = 0;
  let grandTotalRtnQty = 0;

  for (const pc of targetClasses) {
    console.log(`\n==================================================`);
    console.log(`查詢類別: [${pc}]`);
    console.log(`時間範圍: ${begin} ~ ${end}`);
    console.log(`==================================================`);

    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);

    // Filter only "北區" records
    const northRecords = result.records.filter(r => r.zone === "北區");

    if (northRecords.length > 0) {
      console.log(`區域 | 業務 | 經銷 | 產品 | 訂量 | 退量`);
      console.log(`-----|------|------|------|------|------`);
      let classQty = 0;
      let classRtn = 0;
      for (const r of northRecords) {
        console.log(`${r.zone} | ${r.sales} | ${r.customer} | ${r.product} | ${r.qty} | ${r.rtn_qty}`);
        classQty += r.qty;
        classRtn += r.rtn_qty;
      }
      console.log(`\n👉 北區 [${pc}] 小計：訂量 ${classQty} | 退量 ${classRtn}`);
      grandTotalQty += classQty;
      grandTotalRtnQty += classRtn;
    } else {
      console.log(`北區無出貨紀錄。`);
    }
  }

  console.log(`\n==================================================`);
  console.log(`總計 (北區 新思維)：訂量 ${grandTotalQty} | 退量 ${grandTotalRtnQty}`);
  console.log(`==================================================`);
}

run().catch(console.error);
