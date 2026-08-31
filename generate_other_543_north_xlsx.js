// Generate styled XLSX: 114上 其它類別 — 北區各客戶
// Each customer gets their own worksheet showing products grouped by class
// Also includes a 客戶加總 (customer summary) sheet
// Each sheet shows the responsible sales rep name

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_other_543.json'), 'utf-8'));

// =================================================================
// Constants
// =================================================================
const TARGET_ZONE = '北區';
const DATE_RANGE = '114.4.26~114.9.25';

const PRODUCT_CLASSES = [
  { key: '其他:補充類-國文', label: '補充類-國文' },
  { key: '其他:補充類-英語', label: '補充類-英語' },
  { key: '其他:12K喜悅',    label: '12K喜悅' },
  { key: '其他:16K喜悅',    label: '16K喜悅' },
  { key: '其他:其他543',    label: '其他543' },
];

// =================================================================
// Color Palette & Style Definitions
// =================================================================
const COLORS = {
  titleBg:      'FF1F4E79',
  titleFont:    'FFFFFFFF',
  headerBg:     'FF4472C4',
  headerFont:   'FFFFFFFF',
  subHeaderBg:  'FFD6E4F0',
  subHeaderFont:'FF1F4E79',
  sec1Color:    'FF70AD47',  // 補充國文=綠
  sec1LightBg:  'FFE2EFDA',
  sec2Color:    'FF5B9BD5',  // 補充英語=藍
  sec2LightBg:  'FFDAEEF3',
  sec3Color:    'FFED7D31',  // 12K喜悅=橘
  sec3LightBg:  'FFFCE4D6',
  sec4Color:    'FF7030A0',  // 16K喜悅=紫
  sec4LightBg:  'FFE8D5F5',
  sec5Color:    'FF808080',  // 其他543=灰
  sec5LightBg:  'FFF2F2F2',
  totalBg:      'FFFFF2CC',
  totalFont:    'FF833C0B',
  grandTotalBg: 'FFDCE6F1',
  grandTotalFont:'FF1F4E79',
  dataBgEven:   'FFF2F2F2',
  dataBgOdd:    'FFFFFFFF',
  borderColor:  'FF000000',
  negativeBg:   'FFFFC7CE',
  salesRepBg:   'FF2E75B6',  // 業務名稱底色
  salesRepFont: 'FFFFFFCC',
};

const SECTION_COLORS = [
  { color: COLORS.sec1Color, lightBg: COLORS.sec1LightBg },
  { color: COLORS.sec2Color, lightBg: COLORS.sec2LightBg },
  { color: COLORS.sec3Color, lightBg: COLORS.sec3LightBg },
  { color: COLORS.sec4Color, lightBg: COLORS.sec4LightBg },
  { color: COLORS.sec5Color, lightBg: COLORS.sec5LightBg },
];

const FONT_TITLE   = { name: '微軟正黑體', size: 14, bold: true, color: { argb: COLORS.titleFont } };
const FONT_HEADER  = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.headerFont } };
const FONT_SUBHDR  = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.subHeaderFont } };
const FONT_DATA    = { name: '微軟正黑體', size: 10 };
const FONT_BOLD    = { name: '微軟正黑體', size: 10, bold: true };
const FONT_TOTAL   = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.totalFont } };
const FONT_GRAND   = { name: '微軟正黑體', size: 11, bold: true, color: { argb: COLORS.grandTotalFont } };
const FONT_SECTION = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const FONT_DATE    = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFCC' } };
const FONT_SALESREP = { name: '微軟正黑體', size: 11, bold: true, color: { argb: COLORS.salesRepFont } };

const THIN_BORDER = {
  top:    { style: 'thin', color: { argb: COLORS.borderColor } },
  left:   { style: 'thin', color: { argb: COLORS.borderColor } },
  bottom: { style: 'thin', color: { argb: COLORS.borderColor } },
  right:  { style: 'thin', color: { argb: COLORS.borderColor } },
};
const MEDIUM_BORDER_BOTTOM = {
  top:    { style: 'thin',   color: { argb: COLORS.borderColor } },
  left:   { style: 'thin',   color: { argb: COLORS.borderColor } },
  bottom: { style: 'medium', color: { argb: COLORS.borderColor } },
  right:  { style: 'thin',   color: { argb: COLORS.borderColor } },
};

const CENTER = { horizontal: 'center', vertical: 'middle' };
const LEFT_AL = { horizontal: 'left', vertical: 'middle' };
const RIGHT_AL = { horizontal: 'right', vertical: 'middle' };

function solidFill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function styleCell(ws, row, col, opts = {}) {
  const cell = ws.getCell(row, col);
  if (opts.fill) cell.fill = opts.fill;
  if (opts.font) cell.font = opts.font;
  if (opts.border) cell.border = opts.border;
  if (opts.alignment) cell.alignment = opts.alignment;
  if (opts.numFmt) cell.numFmt = opts.numFmt;
  return cell;
}

function chineseSort(a, b) {
  return a.localeCompare(b, 'zh-Hant');
}

function displayName(product) {
  return product.replace(/^\*/, '');
}

// =================================================================
// Data aggregation
// =================================================================
const northRecords = allRecords.filter(r => r.zone === TARGET_ZONE);

// Build customer -> sales rep mapping
const custSalesMap = {};
for (const r of northRecords) {
  if (!custSalesMap[r.customer]) custSalesMap[r.customer] = r.sales;
}

// Build products by class (global across all zones for canonical ordering)
function getProductsByClass(records) {
  const result = {};
  for (const pc of PRODUCT_CLASSES) {
    const classRecs = records.filter(r => r.productClass === pc.key);
    const products = [...new Set(classRecs.map(r => r.product))].sort(chineseSort);
    result[pc.key] = products;
  }
  return result;
}

// Build: customer -> product -> { qty, rtn }
function aggregateByCustomerProduct(records) {
  const data = {};
  for (const r of records) {
    if (!data[r.customer]) data[r.customer] = {};
    if (!data[r.customer][r.product]) data[r.customer][r.product] = { qty: 0, rtn: 0 };
    data[r.customer][r.product].qty += r.qty;
    data[r.customer][r.product].rtn += r.rtn_qty;
  }
  return data;
}

// Build: product -> { qty, rtn } (all customers total)
function aggregateByProduct(records) {
  const data = {};
  for (const r of records) {
    if (!data[r.product]) data[r.product] = { qty: 0, rtn: 0 };
    data[r.product].qty += r.qty;
    data[r.product].rtn += r.rtn_qty;
  }
  return data;
}

const productsByClass = getProductsByClass(allRecords);
const custProductData = aggregateByCustomerProduct(northRecords);
const allProductData = aggregateByProduct(northRecords);
const customers = Object.keys(custSalesMap).sort(chineseSort);

// =================================================================
// Build a customer sheet (or summary sheet)
// =================================================================
function buildCustomerSheet(ws, customerName, salesRepName, productData, productsByClass, isSummary = false) {
  let row = 1;
  const COLS = 6; // 序號, 產品名稱, 類別, 訂量, 退量, 淨出

  // ---- Title Row ----
  const titleText = isSummary
    ? `${TARGET_ZONE} — 其它類別 客戶加總`
    : `${TARGET_ZONE} — 其它類別`;
  ws.mergeCells(row, 1, row, 3);
  ws.getCell(row, 1).value = titleText;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT_AL;

  ws.mergeCells(row, 4, row, COLS);
  ws.getCell(row, 4).value = DATE_RANGE;
  ws.getCell(row, 4).font = FONT_DATE;
  ws.getCell(row, 4).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 4).alignment = RIGHT_AL;

  for (let c = 1; c <= COLS; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) cell.fill = solidFill(COLORS.titleBg);
    cell.border = THIN_BORDER;
  }
  ws.getRow(row).height = 28;
  row++;

  // ---- Customer + Sales Rep Row ----
  ws.mergeCells(row, 1, row, 3);
  ws.getCell(row, 1).value = isSummary ? `全客戶加總 (${customers.length} 客戶)` : `客戶：${customerName}`;
  ws.getCell(row, 1).font = FONT_SUBHDR;
  ws.getCell(row, 1).fill = solidFill(COLORS.subHeaderBg);
  ws.getCell(row, 1).alignment = LEFT_AL;

  ws.mergeCells(row, 4, row, COLS);
  ws.getCell(row, 4).value = isSummary ? `${TARGET_ZONE}` : `業務：${salesRepName}`;
  ws.getCell(row, 4).font = FONT_SALESREP;
  ws.getCell(row, 4).fill = solidFill(COLORS.salesRepBg);
  ws.getCell(row, 4).alignment = RIGHT_AL;

  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
    if (!ws.getCell(row, c).fill || !ws.getCell(row, c).fill.fgColor) {
      ws.getCell(row, c).fill = solidFill(COLORS.subHeaderBg);
    }
  }
  ws.getRow(row).height = 22;
  row++;

  // ---- Header Row ----
  const headers = ['序號', '產品名稱', '類別', '訂量', '退量', '淨出貨'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER,
    });
  }
  row++;

  // ---- Write products grouped by class ----
  let grandQty = 0, grandRtn = 0;
  let hasAnyData = false;

  for (let ci = 0; ci < PRODUCT_CLASSES.length; ci++) {
    const pc = PRODUCT_CLASSES[ci];
    const secColor = SECTION_COLORS[ci];
    const products = productsByClass[pc.key];

    // Determine which products have data for this customer
    const productsWithData = products.filter(p => {
      const d = productData[p];
      return d && (d.qty !== 0 || d.rtn !== 0);
    });

    // For summary sheet, show all products; for customer sheets, only show products with data
    const productsToShow = isSummary ? products : productsWithData;
    if (productsToShow.length === 0 && !isSummary) continue;

    // Section header row
    ws.mergeCells(row, 1, row, COLS);
    ws.getCell(row, 1).value = `  ${pc.label}`;
    ws.getCell(row, 1).font = FONT_SECTION;
    ws.getCell(row, 1).fill = solidFill(secColor.color);
    ws.getCell(row, 1).alignment = LEFT_AL;
    for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
    row++;

    let secQty = 0, secRtn = 0;
    let idx = 1;

    for (const product of productsToShow) {
      const d = productData[product] || { qty: 0, rtn: 0 };
      const net = d.qty - d.rtn;
      // Skip zero-data rows for customer sheets
      if (!isSummary && d.qty === 0 && d.rtn === 0) continue;

      const bgColor = idx % 2 !== 0 ? secColor.lightBg : COLORS.dataBgEven;

      ws.getCell(row, 1).value = idx;
      ws.getCell(row, 2).value = displayName(product);
      ws.getCell(row, 3).value = pc.label;
      ws.getCell(row, 4).value = d.qty || '';
      if (d.qty) ws.getCell(row, 4).numFmt = '#,##0';
      ws.getCell(row, 5).value = d.rtn || '';
      if (d.rtn) ws.getCell(row, 5).numFmt = '#,##0';
      ws.getCell(row, 6).value = net;
      ws.getCell(row, 6).numFmt = '#,##0';

      for (let c = 1; c <= COLS; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(bgColor),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c <= 3 ? (c === 1 ? CENTER : LEFT_AL) : CENTER,
        });
      }

      if (net < 0) {
        ws.getCell(row, 6).fill = solidFill(COLORS.negativeBg);
        ws.getCell(row, 6).font = { ...FONT_DATA, color: { argb: 'FFCC0000' } };
      }

      secQty += d.qty;
      secRtn += d.rtn;
      hasAnyData = true;
      idx++;
      row++;
    }

    // Section subtotal
    const secNet = secQty - secRtn;
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = `${pc.label} 小計`;
    ws.getCell(row, 3).value = '';
    ws.getCell(row, 4).value = secQty || '';
    if (secQty) ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = secRtn || '';
    if (secRtn) ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).value = secNet;
    ws.getCell(row, 6).numFmt = '#,##0';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(secColor.lightBg),
        font: FONT_TOTAL,
        border: MEDIUM_BORDER_BOTTOM,
        alignment: c <= 3 ? LEFT_AL : CENTER,
      });
    }
    ws.getCell(row, 2).font = { ...FONT_TOTAL, color: { argb: secColor.color } };
    row++;

    grandQty += secQty;
    grandRtn += secRtn;

    if (ci < PRODUCT_CLASSES.length - 1) row++; // blank between sections
  }

  // ---- Grand Total ----
  row++;
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '總合計';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = grandQty;
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = grandRtn;
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = grandQty - grandRtn;
  ws.getCell(row, 6).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c <= 3 ? CENTER : CENTER,
    });
  }

  // Column widths
  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 34;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 10;
  ws.getColumn(5).width = 10;
  ws.getColumn(6).width = 10;

  return { qty: grandQty, rtn: grandRtn, hasData: hasAnyData };
}

// =================================================================
// Main
// =================================================================
async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = '金安出版社';
  wb.created = new Date();

  // ---- Sheet 1: 客戶加總 ----
  const summaryWs = wb.addWorksheet('客戶加總', {
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  const summaryResult = buildCustomerSheet(summaryWs, '全客戶', TARGET_ZONE, allProductData, productsByClass, true);
  console.log(`  客戶加總: qty=${summaryResult.qty}, rtn=${summaryResult.rtn}, net=${summaryResult.qty - summaryResult.rtn}`);

  // ---- Per-customer sheets ----
  let custCount = 0;
  let skippedCount = 0;

  for (const customer of customers) {
    const salesRep = custSalesMap[customer];
    const custData = custProductData[customer] || {};

    // Sanitize sheet name (Excel max 31 chars, no special chars)
    let sheetName = customer
      .replace(/[\\/*?:\[\]]/g, '')
      .substring(0, 31);
    // Ensure unique name
    if (wb.worksheets.find(ws => ws.name === sheetName)) {
      sheetName = sheetName.substring(0, 28) + `(${custCount})`;
    }

    const ws = wb.addWorksheet(sheetName, {
      pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    const result = buildCustomerSheet(ws, customer, salesRep, custData, productsByClass, false);
    custCount++;

    if (!result.hasData) {
      skippedCount++;
    }
  }

  const outputDir = path.join(__dirname, 'Output2');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = path.join(outputDir, '114上_其它類別_北區各客戶.xlsx');
  await wb.xlsx.writeFile(outputPath);
  console.log(`\n✓ ${outputPath}`);
  console.log(`  Sheets: ${wb.worksheets.length} (1 summary + ${custCount} customers)`);
  if (skippedCount > 0) {
    console.log(`  Note: ${skippedCount} customers had no data in any category`);
  }

  // Verify: Print per-sales-rep summary
  console.log('\n--- Per Sales Rep Summary ---');
  const salesReps = [...new Set(Object.values(custSalesMap))].sort(chineseSort);
  for (const rep of salesReps) {
    const repCustomers = customers.filter(c => custSalesMap[c] === rep);
    let repQty = 0, repRtn = 0;
    for (const cust of repCustomers) {
      const custData = custProductData[cust] || {};
      for (const vals of Object.values(custData)) {
        repQty += vals.qty;
        repRtn += vals.rtn;
      }
    }
    console.log(`  ${rep}: ${repCustomers.length} customers, qty=${repQty}, rtn=${repRtn}, net=${repQty - repRtn}`);
  }

  console.log('\nDone!');
}

main().catch(console.error);
