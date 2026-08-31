import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  const begin = "2026-5-5";
  const end = "2026-7-17";

  const targetClasses = [
    "國中講義:複習講義(不含5-6)",
    "國中講義:新思維(不含5-6)",
    "國中講義:新思維5-6"
  ];

  const targetSales = ["康晉瑋", "康晉偉"];
  const targetCustomer = "大漢書局";

  const initialRes = await request(BASE_URL, "/sales");
  const parsedInitial = parseSalesPage(initialRes.body);
  if (!parsedInitial.isLoggedIn) {
    console.error("ERROR: Not logged in! Please login first.");
    process.exit(1);
  }
  console.log("Logged in as:", parsedInitial.username);

  const allRecords = [];

  for (const pc of targetClasses) {
    console.log(`Querying [${pc}] from ${begin} to ${end}...`);
    const params = { begin, end, product_class: pc };
    const res = await request(BASE_URL, "/sales", { params });
    const result = parseSalesPage(res.body);

    for (const r of result.records) {
      if (targetSales.includes(r.sales) && r.customer.includes(targetCustomer)) {
        allRecords.push({
          productClass: pc,
          zone: r.zone,
          sales: r.sales,
          customer: r.customer,
          product: r.product,
          qty: r.qty,
          rtn_qty: r.rtn_qty,
        });
      }
    }
  }

  console.log(`Found ${allRecords.length} records for ${targetCustomer} by ${targetSales[0]}.`);

  if (allRecords.length === 0) {
    console.log("No records found.");
    return;
  }

  // Generate Excel
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('出貨統計');
  
  ws.columns = [
    { header: '產品類別', key: 'productClass', width: 25 },
    { header: '業務', key: 'sales', width: 10 },
    { header: '客戶', key: 'customer', width: 20 },
    { header: '產品名稱', key: 'product', width: 40 },
    { header: '訂量', key: 'qty', width: 10 },
    { header: '退量', key: 'rtn_qty', width: 10 },
    { header: '淨出貨', key: 'net', width: 10 },
  ];

  ws.getRow(1).font = { bold: true };

  let totalQty = 0;
  let totalRtn = 0;

  for (const r of allRecords) {
    const net = r.qty - r.rtn_qty;
    totalQty += r.qty;
    totalRtn += r.rtn_qty;
    ws.addRow({
      productClass: r.productClass,
      sales: r.sales,
      customer: r.customer,
      product: r.product,
      qty: r.qty,
      rtn_qty: r.rtn_qty,
      net: net
    });
  }

  ws.addRow({});
  const sumRow = ws.addRow({
    product: '總計',
    qty: totalQty,
    rtn_qty: totalRtn,
    net: totalQty - totalRtn
  });
  sumRow.font = { bold: true };

  const outputDir = path.join(process.cwd(), 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = path.join(outputDir, `康晉偉_大漢書局_複習與新思維出貨統計_20260505_20260717.xlsx`);
  await wb.xlsx.writeFile(outputPath);
  console.log(`Successfully generated: ${outputPath}`);
}

run().catch(console.error);
