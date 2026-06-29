// Generate styled XLSX for 中區 曹原菘 — 複習講義 + 新思維
// 114.4.26~114.9.25
// Structure:
//   Sheet 1: 總表 (summary of all customers)
//   Sheet 2+: One sheet per customer with product details

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_cao_review_xinswei.json'), 'utf-8'));

// =================================================================
// Color Palette & Style Definitions
// =================================================================
const COLORS = {
  titleBg:      'FF1F4E79',  // 深藍
  titleFont:    'FFFFFFFF',
  headerBg:     'FF4472C4',  // 中藍
  headerFont:   'FFFFFFFF',
  subHeaderBg:  'FFD6E4F0',  // 淺藍
  subHeaderFont:'FF1F4E79',
  reviewBg:     'FF70AD47',  // 複習講義=綠
  reviewLightBg:'FFE2EFDA',
  xinswBg:      'FF7030A0',  // 新思維=紫
  xinswLightBg: 'FFE8D5F5',
  totalBg:      'FFFFF2CC',  // 合計=淡黃
  totalFont:    'FF833C0B',
  grandTotalBg: 'FFDCE6F1',  // 總合計=淡藍灰
  dataBgEven:   'FFF2F2F2',
  dataBgOdd:    'FFFFFFFF',
  borderColor:  'FF000000',
  summaryRowBg: 'FFDAEEF3',  // 小計行
  positiveBg:   'FFE2EFDA',  // 正數淨出=淡綠
  negativeBg:   'FFFCE4D6',  // 負數淨出=淡紅
};

const FONT_TITLE = { name: '微軟正黑體', size: 14, bold: true, color: { argb: COLORS.titleFont } };
const FONT_HEADER = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.headerFont } };
const FONT_SUBHEADER = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.subHeaderFont } };
const FONT_DATA = { name: '微軟正黑體', size: 10 };
const FONT_TOTAL = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.totalFont } };
const FONT_GRAND = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FF1F4E79' } };
const FONT_SECTION = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const FONT_DATE = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFCC' } };

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
const LEFT = { horizontal: 'left', vertical: 'middle' };
const RIGHT = { horizontal: 'right', vertical: 'middle' };

function solidFill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function styleCell(ws, row, col, { fill, font, border, alignment, numFmt } = {}) {
  const cell = ws.getCell(row, col);
  if (fill) cell.fill = fill;
  if (font) cell.font = font;
  if (border) cell.border = border;
  if (alignment) cell.alignment = alignment;
  if (numFmt) cell.numFmt = numFmt;
}

function styleRange(ws, r1, c1, r2, c2, styles) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      styleCell(ws, r, c, styles);
    }
  }
}

// =================================================================
// Data aggregation
// =================================================================
function aggregateData(records) {
  // Group by customer -> product, sum qty/rtn
  const byCustomer = {};
  for (const r of records) {
    if (!byCustomer[r.customer]) byCustomer[r.customer] = {};
    const key = r.product;
    if (!byCustomer[r.customer][key]) {
      byCustomer[r.customer][key] = { productClass: r.productClass, qty: 0, rtn: 0 };
    }
    byCustomer[r.customer][key].qty += r.qty;
    byCustomer[r.customer][key].rtn += r.rtn_qty;
  }
  return byCustomer;
}

// Define the canonical product order for display
const REVIEW_PRODUCTS = [
  '*雙向複習 國文(全)',
  '*雙向複習 國文(1-4)',
  '*雙向複習 英語(全)',
  '*雙向複習 英語(1-4)',
  '*雙向複習 數學(全)',
  '*雙向複習 數學(1-4)',
  '*雙向複習 理化(全)',
  '*雙向複習 理化(3-4)',
  '*雙向複習 生物(全)',
  '*雙向複習 地球科學(全)',
  '*雙向複習 地理(全)',
  '*雙向複習 歷史(全)',
  '*雙向複習 公民(全)',
  '*735輕鬆讀 圖解複習 英語(1-4)',
  '*735輕鬆讀 圖解複習 數學(1-4)',
  '*735輕鬆讀 圖解複習 理化(3-4)',
  '*735輕鬆讀 圖解複習 公民(全)',
  '*(主題讚)複習 國文(全)',
];

const XINSWEI_PRODUCTS = [
  '*新思維複習 國文(1-4)',
  '*新思維複習 英語(1-4)',
];

// Group labels for display
function getProductGroup(product) {
  if (product.startsWith('*雙向複習')) return '雙向';
  if (product.startsWith('*735')) return '735';
  if (product.startsWith('*(主題讚)')) return '主題讚';
  if (product.startsWith('*新思維')) return '新思維';
  return '其他';
}

// Short display name for products
function shortName(product) {
  return product
    .replace(/^\*/, '')
    .replace('735輕鬆讀 圖解複習', '735複習')
    .replace('(主題讚)複習', '主題讚複習');
}

// =================================================================
// Build summary sheet (總表)
// =================================================================
function buildSummarySheet(ws, byCustomer) {
  const customers = Object.keys(byCustomer).sort();
  const allProducts = [...REVIEW_PRODUCTS, ...XINSWEI_PRODUCTS];
  const TOTAL_COLS = 3 + allProducts.length + 3; // No + 經銷 + 產品columns... 
  // Columns: A=序號, B=經銷, C..=products, then 訂書合計, 退貨合計, 淨出合計

  let row = 1;
  const lastCol = 2 + allProducts.length + 3; // 2 fixed + products + 3 summary

  // ---- Title Row ----
  ws.mergeCells(row, 1, row, Math.floor(lastCol / 2));
  ws.getCell(row, 1).value = '114上 複習講義+新思維 出貨統計 — 曹原菘';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = CENTER;

  ws.mergeCells(row, Math.floor(lastCol / 2) + 1, row, lastCol);
  ws.getCell(row, Math.floor(lastCol / 2) + 1).value = '114.4.26~114.9.25';
  ws.getCell(row, Math.floor(lastCol / 2) + 1).font = FONT_DATE;
  ws.getCell(row, Math.floor(lastCol / 2) + 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, Math.floor(lastCol / 2) + 1).alignment = { horizontal: 'right', vertical: 'middle' };

  // Fill the rest of the title row
  for (let c = 1; c <= lastCol; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) cell.fill = solidFill(COLORS.titleBg);
    cell.border = THIN_BORDER;
  }
  row++;

  // ---- Category Header Row ----
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '';

  // Find where review products end and xinswei starts
  const reviewEnd = 2 + REVIEW_PRODUCTS.length; // last review col
  const xinswStart = reviewEnd + 1;
  const xinswEnd = xinswStart + XINSWEI_PRODUCTS.length - 1;

  if (REVIEW_PRODUCTS.length > 0) {
    ws.mergeCells(row, 3, row, 2 + REVIEW_PRODUCTS.length);
    ws.getCell(row, 3).value = '複習講義';
    ws.getCell(row, 3).font = FONT_SECTION;
    ws.getCell(row, 3).fill = solidFill(COLORS.reviewBg);
    ws.getCell(row, 3).alignment = CENTER;
  }
  if (XINSWEI_PRODUCTS.length > 0) {
    ws.mergeCells(row, xinswStart, row, xinswEnd);
    ws.getCell(row, xinswStart).value = '新思維';
    ws.getCell(row, xinswStart).font = FONT_SECTION;
    ws.getCell(row, xinswStart).fill = solidFill(COLORS.xinswBg);
    ws.getCell(row, xinswStart).alignment = CENTER;
  }

  // Summary header cells
  const summCol = xinswEnd + 1;
  ws.getCell(row, summCol).value = '';
  ws.getCell(row, summCol + 1).value = '';
  ws.getCell(row, summCol + 2).value = '';

  for (let c = 1; c <= lastCol; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.border) cell.border = THIN_BORDER;
    if (!cell.fill || !cell.fill.fgColor) {
      cell.fill = solidFill(COLORS.subHeaderBg);
      cell.font = FONT_SUBHEADER;
    }
    cell.border = THIN_BORDER;
  }
  row++;

  // ---- Product Header Row ----
  ws.getCell(row, 1).value = '序號';
  ws.getCell(row, 2).value = '經銷';
  
  let col = 3;
  for (const p of allProducts) {
    ws.getCell(row, col).value = shortName(p);
    ws.getCell(row, col).alignment = { ...CENTER, wrapText: true };
    col++;
  }
  ws.getCell(row, col).value = '訂書合計';
  ws.getCell(row, col + 1).value = '退貨合計';
  ws.getCell(row, col + 2).value = '淨出合計';

  for (let c = 1; c <= lastCol; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: { ...CENTER, wrapText: true },
    });
  }
  ws.getRow(row).height = 48;
  row++;

  // ---- Data Rows ----
  let idx = 1;
  const grandTotals = { qty: 0, rtn: 0 };
  const productGrandTotals = {};

  for (const customer of customers) {
    const custData = byCustomer[customer];
    const bgColor = idx % 2 === 0 ? COLORS.dataBgEven : COLORS.dataBgOdd;

    ws.getCell(row, 1).value = idx;
    ws.getCell(row, 2).value = customer;
    
    let custQty = 0;
    let custRtn = 0;
    col = 3;
    for (const p of allProducts) {
      const d = custData[p];
      const qty = d ? d.qty : 0;
      const rtn = d ? d.rtn : 0;
      const net = qty - rtn;
      ws.getCell(row, col).value = net !== 0 ? net : '';
      if (net !== 0) ws.getCell(row, col).numFmt = '#,##0';
      custQty += qty;
      custRtn += rtn;
      productGrandTotals[p] = (productGrandTotals[p] || 0) + net;
      col++;
    }

    ws.getCell(row, col).value = custQty;
    ws.getCell(row, col).numFmt = '#,##0';
    ws.getCell(row, col + 1).value = custRtn;
    ws.getCell(row, col + 1).numFmt = '#,##0';
    ws.getCell(row, col + 2).value = custQty - custRtn;
    ws.getCell(row, col + 2).numFmt = '#,##0';

    grandTotals.qty += custQty;
    grandTotals.rtn += custRtn;

    // Styling
    for (let c = 1; c <= lastCol; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER,
        alignment: c <= 2 ? LEFT : CENTER,
      });
    }
    // Highlight summary columns
    for (let sc = col; sc <= col + 2; sc++) {
      ws.getCell(row, sc).fill = solidFill(COLORS.summaryRowBg);
      ws.getCell(row, sc).font = { ...FONT_DATA, bold: true };
    }

    row++;
    idx++;
  }

  // ---- Grand Total Row ----
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '合計';
  col = 3;
  for (const p of allProducts) {
    const net = productGrandTotals[p] || 0;
    ws.getCell(row, col).value = net !== 0 ? net : '';
    if (net !== 0) ws.getCell(row, col).numFmt = '#,##0';
    col++;
  }
  ws.getCell(row, col).value = grandTotals.qty;
  ws.getCell(row, col).numFmt = '#,##0';
  ws.getCell(row, col + 1).value = grandTotals.rtn;
  ws.getCell(row, col + 1).numFmt = '#,##0';
  ws.getCell(row, col + 2).value = grandTotals.qty - grandTotals.rtn;
  ws.getCell(row, col + 2).numFmt = '#,##0';

  for (let c = 1; c <= lastCol; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c <= 2 ? CENTER : CENTER,
    });
  }

  // ---- Column widths ----
  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 16;
  for (let c = 3; c <= 2 + allProducts.length; c++) {
    ws.getColumn(c).width = 9;
  }
  const sc = 2 + allProducts.length + 1;
  ws.getColumn(sc).width = 10;
  ws.getColumn(sc + 1).width = 10;
  ws.getColumn(sc + 2).width = 10;

  ws.getRow(1).height = 28;
}

// =================================================================
// Build customer detail sheet
// =================================================================
function buildCustomerSheet(ws, customer, custData) {
  let row = 1;
  const COLS = 6; // 產品名稱, 類別, 訂量, 退量, 淨出, 退書率

  // ---- Title ----
  ws.mergeCells(row, 1, row, 3);
  ws.getCell(row, 1).value = `${customer} — 複習講義+新思維`;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 4, row, COLS);
  ws.getCell(row, 4).value = '114.4.26~114.9.25';
  ws.getCell(row, 4).font = FONT_DATE;
  ws.getCell(row, 4).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };

  for (let c = 1; c <= COLS; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) cell.fill = solidFill(COLORS.titleBg);
    cell.border = THIN_BORDER;
  }
  ws.getRow(row).height = 28;
  row++;

  // ---- Header ----
  const headers = ['產品名稱', '類別', '訂量', '退量', '淨出貨', '退書率'];
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

  // ---- Section: 複習講義 ----
  const reviewProducts = REVIEW_PRODUCTS.filter(p => custData[p]);
  const xinswProducts = XINSWEI_PRODUCTS.filter(p => custData[p]);

  function writeSection(sectionName, products, sectionColor, sectionLightBg) {
    if (products.length === 0) return;
    
    // Section header
    ws.mergeCells(row, 1, row, COLS);
    ws.getCell(row, 1).value = `  ${sectionName}`;
    ws.getCell(row, 1).font = FONT_SECTION;
    ws.getCell(row, 1).fill = solidFill(sectionColor);
    ws.getCell(row, 1).alignment = LEFT;
    ws.getCell(row, 1).border = THIN_BORDER;
    for (let c = 2; c <= COLS; c++) {
      ws.getCell(row, c).border = THIN_BORDER;
    }
    row++;

    let sectionQty = 0;
    let sectionRtn = 0;
    let rowIdx = 0;

    for (const p of products) {
      const d = custData[p];
      const qty = d.qty;
      const rtn = d.rtn;
      const net = qty - rtn;
      const rtnRate = qty > 0 ? (rtn / qty * 100) : 0;
      const bgColor = rowIdx % 2 === 0 ? sectionLightBg : COLORS.dataBgOdd;

      ws.getCell(row, 1).value = shortName(p);
      ws.getCell(row, 2).value = getProductGroup(p);
      ws.getCell(row, 3).value = qty;
      ws.getCell(row, 3).numFmt = '#,##0';
      ws.getCell(row, 4).value = rtn;
      ws.getCell(row, 4).numFmt = '#,##0';
      ws.getCell(row, 5).value = net;
      ws.getCell(row, 5).numFmt = '#,##0';
      ws.getCell(row, 6).value = rtnRate > 0 ? rtnRate / 100 : '';
      if (rtnRate > 0) ws.getCell(row, 6).numFmt = '0.0%';

      for (let c = 1; c <= COLS; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(bgColor),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c <= 2 ? LEFT : CENTER,
        });
      }

      // Highlight negative net
      if (net < 0) {
        ws.getCell(row, 5).fill = solidFill(COLORS.negativeBg);
        ws.getCell(row, 5).font = { ...FONT_DATA, color: { argb: 'FFCC0000' } };
      }

      sectionQty += qty;
      sectionRtn += rtn;
      rowIdx++;
      row++;
    }

    // Section subtotal
    const sectionNet = sectionQty - sectionRtn;
    const sectionRate = sectionQty > 0 ? (sectionRtn / sectionQty * 100) : 0;
    
    ws.getCell(row, 1).value = `${sectionName} 小計`;
    ws.getCell(row, 2).value = '';
    ws.getCell(row, 3).value = sectionQty;
    ws.getCell(row, 3).numFmt = '#,##0';
    ws.getCell(row, 4).value = sectionRtn;
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = sectionNet;
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).value = sectionRate > 0 ? sectionRate / 100 : '';
    if (sectionRate > 0) ws.getCell(row, 6).numFmt = '0.0%';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(COLORS.summaryRowBg),
        font: FONT_TOTAL,
        border: MEDIUM_BORDER_BOTTOM,
        alignment: c <= 2 ? LEFT : CENTER,
      });
    }
    row++;
    
    return { qty: sectionQty, rtn: sectionRtn };
  }

  const reviewTotals = writeSection('複習講義', reviewProducts, COLORS.reviewBg, COLORS.reviewLightBg) || { qty: 0, rtn: 0 };
  
  // Blank separator
  row++;
  
  const xinswTotals = writeSection('新思維', xinswProducts, COLORS.xinswBg, COLORS.xinswLightBg) || { qty: 0, rtn: 0 };

  // ---- Grand Total ----
  row++;
  const grandQty = reviewTotals.qty + xinswTotals.qty;
  const grandRtn = reviewTotals.rtn + xinswTotals.rtn;
  const grandNet = grandQty - grandRtn;
  const grandRate = grandQty > 0 ? (grandRtn / grandQty * 100) : 0;

  ws.getCell(row, 1).value = '總合計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = grandQty;
  ws.getCell(row, 3).numFmt = '#,##0';
  ws.getCell(row, 4).value = grandRtn;
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = grandNet;
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = grandRate > 0 ? grandRate / 100 : '';
  if (grandRate > 0) ws.getCell(row, 6).numFmt = '0.0%';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c <= 2 ? CENTER : CENTER,
    });
  }

  // Column widths
  ws.getColumn(1).width = 30;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 10;
  ws.getColumn(4).width = 10;
  ws.getColumn(5).width = 10;
  ws.getColumn(6).width = 10;
}

// =================================================================
// Main
// =================================================================
async function main() {
  const byCustomer = aggregateData(allRecords);
  const customers = Object.keys(byCustomer).sort();

  const wb = new ExcelJS.Workbook();
  wb.creator = '金安出版社';
  wb.created = new Date();

  // Sheet 1: 總表
  const summaryWs = wb.addWorksheet('總表', {
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  buildSummarySheet(summaryWs, byCustomer);

  // Sheet 2+: Per-customer
  for (const customer of customers) {
    let sheetName = customer.replace(/[\\\/\?\*\[\]]/g, '');
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

    const ws = wb.addWorksheet(sheetName, {
      pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    buildCustomerSheet(ws, customer, byCustomer[customer]);
  }

  const outputDir = path.join(__dirname, 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = path.join(outputDir, '曹原菘_114上複習講義_新思維.xlsx');
  await wb.xlsx.writeFile(outputPath);
  console.log(`✓ ${outputPath} (${customers.length + 1} sheets: 總表 + ${customers.length} customers)`);
  console.log('Done!');
}

main().catch(console.error);
