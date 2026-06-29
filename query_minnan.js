import { request, loadCookies } from './dist/client.js';
import { parseSalesPage, formatSalesResults } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const dateRanges = [
    { begin: "2026-5-29", end: "2026-5-29", label: "2026-05-29 單日" },
    { begin: "2026-5-26", end: "2026-5-29", label: "2026-05-26 至 2026-05-29 區間" },
    { begin: "2026-5-1", end: "2026-5-31", label: "2026年5月份整月" }
  ];
  const classes = [
    "本土語:國小閩語-課本",
    "本土語:國小閩語-CD",
    "本土語:國小閩語-備課"
  ];

  for (const range of dateRanges) {
    console.log(`\n==================================================`);
    console.log(`🔍 查詢時間範圍: ${range.label} (${range.begin} ~ ${range.end})`);
    console.log(`==================================================`);
    
    for (const pc of classes) {
      const params = {
        begin: range.begin,
        end: range.end,
        product_class: pc
      };
      try {
        const res = await request(BASE_URL, "/sales", { params });
        const result = parseSalesPage(res.body);
        result.period = `${range.begin} ~ ${range.end}`;
        result.productClass = pc;
        if (result.records.length > 0) {
          console.log(`\n📁 類別 [${pc}] 有資料：`);
          console.log(formatSalesResults(result));
        } else {
          console.log(`📁 類別 [${pc}]: （查無資料）`);
        }
      } catch (e) {
        console.error(`查詢失敗 [${pc}]:`, e);
      }
    }
  }
}

run().catch(console.error);
