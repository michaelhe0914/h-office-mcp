import { request, loadCookies } from './dist/client.js';
import { parseSalesPage, formatSalesResults } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const params = {
    begin: "2026-5-15",
    end: "2026-7-23",
    product_class: "國小:國小簿本-彩封"
  };

  console.log(`Querying ${params.begin} to ${params.end} for [${params.product_class}]...`);
  const res = await request(BASE_URL, "/sales", { params });
  const result = parseSalesPage(res.body);

  console.log("=== RECORD SAMPLE ===");
  if (result.records.length > 0) {
    console.log(JSON.stringify(result.records[0], null, 2));
  }

  // Find all unique zones and sales reps in records
  const zones = new Set();
  const salesReps = new Set();
  result.records.forEach(r => {
    if (r.zone) zones.add(r.zone);
    if (r.sales) salesReps.add(r.sales);
  });
  console.log("\n=== UNIQUE ZONES ===");
  console.log(Array.from(zones));
  console.log("=== UNIQUE SALES REPS ===");
  console.log(Array.from(salesReps));

  // Filter North Region records (zone === '北區')
  const northRecords = result.records.filter(r => r.zone === '北區');
  console.log(`\nTotal North Region Records: ${northRecords.length}`);

  let northTotalQty = 0;
  let northTotalRtnQty = 0;

  // Breakdown by sales rep
  const repSummary = {};
  // Breakdown by product
  const productSummary = {};
  // Breakdown by customer
  const customerSummary = {};

  northRecords.forEach(r => {
    const qty = Number(r.qty || 0);
    const rtnQty = Number(r.rtn_qty || 0);
    const netQty = qty - rtnQty;

    northTotalQty += qty;
    northTotalRtnQty += rtnQty;

    const rep = r.sales || "未知業務";
    if (!repSummary[rep]) {
      repSummary[rep] = { shipQty: 0, returnQty: 0, netQty: 0, recordCount: 0 };
    }
    repSummary[rep].shipQty += qty;
    repSummary[rep].returnQty += rtnQty;
    repSummary[rep].netQty += netQty;
    repSummary[rep].recordCount += 1;

    const prod = r.product || "未知品項";
    if (!productSummary[prod]) {
      productSummary[prod] = { shipQty: 0, returnQty: 0, netQty: 0 };
    }
    productSummary[prod].shipQty += qty;
    productSummary[prod].returnQty += rtnQty;
    productSummary[prod].netQty += netQty;

    const cust = r.customer || "未知客戶";
    if (!customerSummary[cust]) {
      customerSummary[cust] = { shipQty: 0, returnQty: 0, netQty: 0 };
    }
    customerSummary[cust].shipQty += qty;
    customerSummary[cust].returnQty += rtnQty;
    customerSummary[cust].netQty += netQty;
  });

  const northNetQty = northTotalQty - northTotalRtnQty;

  console.log("\n==========================================");
  console.log("2026/5/15 ~ 7/23 北區國小彩封數量 統計結果");
  console.log("==========================================");
  console.log(`總出貨數量 (Ship Qty)   : ${northTotalQty.toLocaleString()}`);
  console.log(`總退貨數量 (Return Qty) : ${northTotalRtnQty.toLocaleString()}`);
  console.log(`總淨出貨量 (Net Qty)    : ${northNetQty.toLocaleString()}`);

  console.log("\n--- 各業務統計 (Sales Rep Breakdown) ---");
  console.table(Object.entries(repSummary).map(([sales, s]) => ({
    業務: sales,
    出貨量: s.shipQty.toLocaleString(),
    退貨量: s.returnQty.toLocaleString(),
    淨出貨量: s.netQty.toLocaleString(),
    筆數: s.recordCount
  })));

  console.log("\n--- 品項統計 (Product Summary - Top 20) ---");
  const sortedProds = Object.entries(productSummary)
    .sort((a, b) => b[1].netQty - a[1].netQty);
  console.table(sortedProds.slice(0, 20).map(([prod, s]) => ({
    品項名稱: prod,
    出貨量: s.shipQty.toLocaleString(),
    退貨量: s.returnQty.toLocaleString(),
    淨出貨量: s.netQty.toLocaleString()
  })));
  console.log("\n--- 所有品項統計明細 (All 34 Products) ---");
  console.table(sortedProds.map(([prod, s]) => ({
    品項名稱: prod,
    出貨量: s.shipQty.toLocaleString(),
    退貨量: s.returnQty.toLocaleString(),
    淨出貨量: s.netQty.toLocaleString()
  })));

  console.log("\n--- 各業務對應客戶銷售統計 ---");
  const repCustomerSummary = {};
  northRecords.forEach(r => {
    const rep = r.sales || "未知業務";
    const cust = r.customer || "未知客戶";
    const qty = Number(r.qty || 0);
    const rtnQty = Number(r.rtn_qty || 0);
    const key = `${rep} - ${cust}`;
    if (!repCustomerSummary[key]) {
      repCustomerSummary[key] = { sales: rep, customer: cust, shipQty: 0, returnQty: 0, netQty: 0 };
    }
    repCustomerSummary[key].shipQty += qty;
    repCustomerSummary[key].returnQty += rtnQty;
    repCustomerSummary[key].netQty += qty - rtnQty;
  });
  console.table(Object.values(repCustomerSummary));
}

run().catch(console.error);





