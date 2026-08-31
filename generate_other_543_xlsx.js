// Generate styled XLSX: 114上 其它類別 — 北區、中區、南區
// 列出各類別底下的個別產品
// Structure:
//   Sheet 1: 總表 (all zones summary, each product grouped by class)
//   Sheet 2: 北區 (per-product detail grouped by class)
//   Sheet 3: 中區
//   Sheet 4: 南區

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_other_543.json'), 'utf-8'));

// =================================================================
// Constants
// =================================================================
const TARGET_ZONES = ['北區', '中區', '南區'];
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
  // Section colors for each product class
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
  // Summary
  totalBg:      'FFFFF2CC',
  totalFont:    'FF833C0B',
  grandTotalBg: 'FFDCE6F1',
  grandTotalFont:'FF1F4E79',
  // Data rows
  dataBgEven:   'FFF2F2F2',
  dataBgOdd:    'FFFFFFFF',
  borderColor:  'FF000000',
  // Zone colors
  northBg:      'FFD6E4F0',
  centralBg:    'FFE2EFDA',
  southBg:      'FFFCE4D6',
  negativeBg:   'FFFFC7CE',
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

// Clean product name for display (remove leading *)
function displayName(product) {
  return product.replace(/^\*/, '');
}

// =================================================================
// Data aggregation helpers
// =================================================================

// Get unique products per class across all zones, sorted
function getProductsByClass(records) {
  const result = {};
  for (const pc of PRODUCT_CLASSES) {
    const classRecs = records.filter(r => r.productClass === pc.key);
    const products = [...new Set(classRecs.map(r => r.product))].sort(chineseSort);
    result[pc.key] = products;
  }
  return result;
}

// Aggregate: zone -> product -> { qty, rtn }
function aggregateByZoneProduct(records) {
  const data = {};
  for (const r of records) {
    if (!TARGET_ZONES.includes(r.zone)) continue;
    if (!data[r.zone]) data[r.zone] = {};
    const key = r.product;
    if (!data[r.zone][key]) data[r.zone][key] = { qty: 0, rtn: 0 };
    data[r.zone][key].qty += r.qty;
    data[r.zone][key].rtn += r.rtn_qty;
  }
  return data;
}

// =================================================================
// Build 總表 sheet — products grouped by class, with zone columns
// =================================================================
function buildSummarySheet(ws, zoneProductData, productsByClass) {
  let row = 1;
  // Columns: A=序號, B=產品名稱,
  // then for each zone: 訂量, 退量, 淨出 (3 cols each)
  // then: 全區合計 訂量, 退量, 淨出
  const zoneColStart = 3;
  const colsPerZone = 3;
  const totalZoneCols = TARGET_ZONES.length * colsPerZone;
  const summaryStart = zoneColStart + totalZoneCols;
  const lastCol = summaryStart + 2;

  // ---- Title Row ----
  ws.mergeCells(row, 1, row, Math.floor(lastCol / 2));
  ws.getCell(row, 1).value = '114上 其它類別各區產品數量統計';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = CENTER;

  ws.mergeCells(row, Math.floor(lastCol / 2) + 1, row, lastCol);
  ws.getCell(row, Math.floor(lastCol / 2) + 1).value = DATE_RANGE;
  ws.getCell(row, Math.floor(lastCol / 2) + 1).font = FONT_DATE;
  ws.getCell(row, Math.floor(lastCol / 2) + 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, Math.floor(lastCol / 2) + 1).alignment = RIGHT_AL;

  for (let c = 1; c <= lastCol; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) cell.fill = solidFill(COLORS.titleBg);
    cell.border = THIN_BORDER;
  }
  ws.getRow(row).height = 28;
  row++;

  // ---- Zone Header Row ----
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '';
  for (let c = 1; c <= 2; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.subHeaderBg);
    ws.getCell(row, c).border = THIN_BORDER;
  }

  const zoneBgs = [COLORS.northBg, COLORS.centralBg, COLORS.southBg];
  let col = zoneColStart;
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    ws.mergeCells(row, col, row, col + 2);
    ws.getCell(row, col).value = TARGET_ZONES[i];
    ws.getCell(row, col).font = FONT_SUBHDR;
    ws.getCell(row, col).fill = solidFill(zoneBgs[i]);
    ws.getCell(row, col).alignment = CENTER;
    for (let j = 0; j < 3; j++) ws.getCell(row, col + j).border = THIN_BORDER;
    col += 3;
  }

  ws.mergeCells(row, summaryStart, row, summaryStart + 2);
  ws.getCell(row, summaryStart).value = '全區合計';
  ws.getCell(row, summaryStart).font = FONT_SUBHDR;
  ws.getCell(row, summaryStart).fill = solidFill(COLORS.grandTotalBg);
  ws.getCell(row, summaryStart).alignment = CENTER;
  for (let j = 0; j < 3; j++) ws.getCell(row, summaryStart + j).border = THIN_BORDER;
  row++;

  // ---- Sub-header Row ----
  ws.getCell(row, 1).value = '序號';
  ws.getCell(row, 2).value = '產品名稱';
  for (let c = 1; c <= 2; c++) {
    ws.getCell(row, c).font = FONT_HEADER;
    ws.getCell(row, c).fill = solidFill(COLORS.headerBg);
    ws.getCell(row, c).alignment = CENTER;
    ws.getCell(row, c).border = THIN_BORDER;
  }

  col = zoneColStart;
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    for (let j = 0; j < 3; j++) {
      ws.getCell(row, col + j).value = ['訂量', '退量', '淨出'][j];
      ws.getCell(row, col + j).fill = solidFill(zoneBgs[i]);
      ws.getCell(row, col + j).font = FONT_SUBHDR;
      ws.getCell(row, col + j).alignment = CENTER;
      ws.getCell(row, col + j).border = THIN_BORDER;
    }
    col += 3;
  }
  for (let j = 0; j < 3; j++) {
    ws.getCell(row, summaryStart + j).value = ['訂量', '退量', '淨出'][j];
    ws.getCell(row, summaryStart + j).fill = solidFill(COLORS.grandTotalBg);
    ws.getCell(row, summaryStart + j).font = FONT_SUBHDR;
    ws.getCell(row, summaryStart + j).alignment = CENTER;
    ws.getCell(row, summaryStart + j).border = THIN_BORDER;
  }
  row++;

  // ---- Data Rows grouped by class ----
  let grandTotals = TARGET_ZONES.map(() => ({ qty: 0, rtn: 0 }));
  let allQty = 0, allRtn = 0;

  for (let ci = 0; ci < PRODUCT_CLASSES.length; ci++) {
    const pc = PRODUCT_CLASSES[ci];
    const secColor = SECTION_COLORS[ci];
    const products = productsByClass[pc.key];

    // Section header row
    ws.mergeCells(row, 1, row, lastCol);
    ws.getCell(row, 1).value = `  ${pc.label}`;
    ws.getCell(row, 1).font = FONT_SECTION;
    ws.getCell(row, 1).fill = solidFill(secColor.color);
    ws.getCell(row, 1).alignment = LEFT_AL;
    for (let c = 1; c <= lastCol; c++) ws.getCell(row, c).border = THIN_BORDER;
    row++;

    let secTotals = TARGET_ZONES.map(() => ({ qty: 0, rtn: 0 }));
    let secAllQty = 0, secAllRtn = 0;
    let idx = 1;

    for (const product of products) {
      const bgColor = idx % 2 === 0 ? COLORS.dataBgEven : secColor.lightBg;

      ws.getCell(row, 1).value = idx;
      ws.getCell(row, 2).value = displayName(product);
      ws.getCell(row, 1).font = FONT_DATA;
      ws.getCell(row, 1).fill = solidFill(bgColor);
      ws.getCell(row, 1).alignment = CENTER;
      ws.getCell(row, 1).border = THIN_BORDER;
      ws.getCell(row, 2).font = FONT_DATA;
      ws.getCell(row, 2).fill = solidFill(bgColor);
      ws.getCell(row, 2).alignment = LEFT_AL;
      ws.getCell(row, 2).border = THIN_BORDER;

      col = zoneColStart;
      let rowQty = 0, rowRtn = 0;
      for (let i = 0; i < TARGET_ZONES.length; i++) {
        const zone = TARGET_ZONES[i];
        const d = zoneProductData[zone]?.[product] || { qty: 0, rtn: 0 };
        const net = d.qty - d.rtn;
        ws.getCell(row, col).value = d.qty || '';
        ws.getCell(row, col + 1).value = d.rtn || '';
        ws.getCell(row, col + 2).value = net || '';

        for (let j = 0; j < 3; j++) {
          const cell = ws.getCell(row, col + j);
          cell.numFmt = '#,##0';
          cell.fill = solidFill(bgColor);
          cell.font = FONT_DATA;
          cell.alignment = CENTER;
          cell.border = THIN_BORDER;
        }
        // Negative highlighting
        if (net < 0) {
          ws.getCell(row, col + 2).fill = solidFill(COLORS.negativeBg);
          ws.getCell(row, col + 2).font = { ...FONT_DATA, color: { argb: 'FFCC0000' } };
        }

        secTotals[i].qty += d.qty;
        secTotals[i].rtn += d.rtn;
        rowQty += d.qty;
        rowRtn += d.rtn;
        col += 3;
      }

      // Row total
      const rowNet = rowQty - rowRtn;
      ws.getCell(row, summaryStart).value = rowQty || '';
      ws.getCell(row, summaryStart + 1).value = rowRtn || '';
      ws.getCell(row, summaryStart + 2).value = rowNet || '';
      for (let j = 0; j < 3; j++) {
        const cell = ws.getCell(row, summaryStart + j);
        cell.numFmt = '#,##0';
        cell.fill = solidFill(idx % 2 === 0 ? 'FFEDF2F9' : 'FFF5F8FC');
        cell.font = FONT_BOLD;
        cell.alignment = CENTER;
        cell.border = THIN_BORDER;
      }
      if (rowNet < 0) {
        ws.getCell(row, summaryStart + 2).fill = solidFill(COLORS.negativeBg);
        ws.getCell(row, summaryStart + 2).font = { ...FONT_BOLD, color: { argb: 'FFCC0000' } };
      }

      secAllQty += rowQty;
      secAllRtn += rowRtn;
      idx++;
      row++;
    }

    // Section subtotal row
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = `${pc.label} 小計`;
    ws.getCell(row, 1).fill = solidFill(secColor.lightBg);
    ws.getCell(row, 1).border = MEDIUM_BORDER_BOTTOM;
    ws.getCell(row, 2).font = { ...FONT_TOTAL, color: { argb: secColor.color } };
    ws.getCell(row, 2).fill = solidFill(secColor.lightBg);
    ws.getCell(row, 2).alignment = LEFT_AL;
    ws.getCell(row, 2).border = MEDIUM_BORDER_BOTTOM;

    col = zoneColStart;
    for (let i = 0; i < TARGET_ZONES.length; i++) {
      ws.getCell(row, col).value = secTotals[i].qty;
      ws.getCell(row, col + 1).value = secTotals[i].rtn;
      ws.getCell(row, col + 2).value = secTotals[i].qty - secTotals[i].rtn;
      for (let j = 0; j < 3; j++) {
        const cell = ws.getCell(row, col + j);
        cell.numFmt = '#,##0';
        cell.fill = solidFill(secColor.lightBg);
        cell.font = FONT_TOTAL;
        cell.alignment = CENTER;
        cell.border = MEDIUM_BORDER_BOTTOM;
      }
      grandTotals[i].qty += secTotals[i].qty;
      grandTotals[i].rtn += secTotals[i].rtn;
      col += 3;
    }
    ws.getCell(row, summaryStart).value = secAllQty;
    ws.getCell(row, summaryStart + 1).value = secAllRtn;
    ws.getCell(row, summaryStart + 2).value = secAllQty - secAllRtn;
    for (let j = 0; j < 3; j++) {
      const cell = ws.getCell(row, summaryStart + j);
      cell.numFmt = '#,##0';
      cell.fill = solidFill(secColor.lightBg);
      cell.font = FONT_TOTAL;
      cell.alignment = CENTER;
      cell.border = MEDIUM_BORDER_BOTTOM;
    }

    allQty += secAllQty;
    allRtn += secAllRtn;
    row++;
  }

  // ---- Grand Total Row ----
  row++;
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '總合計';
  ws.getCell(row, 1).fill = solidFill(COLORS.totalBg);
  ws.getCell(row, 1).border = MEDIUM_BORDER_BOTTOM;
  ws.getCell(row, 2).font = FONT_GRAND;
  ws.getCell(row, 2).fill = solidFill(COLORS.totalBg);
  ws.getCell(row, 2).alignment = CENTER;
  ws.getCell(row, 2).border = MEDIUM_BORDER_BOTTOM;

  col = zoneColStart;
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    ws.getCell(row, col).value = grandTotals[i].qty;
    ws.getCell(row, col + 1).value = grandTotals[i].rtn;
    ws.getCell(row, col + 2).value = grandTotals[i].qty - grandTotals[i].rtn;
    for (let j = 0; j < 3; j++) {
      const cell = ws.getCell(row, col + j);
      cell.numFmt = '#,##0';
      cell.fill = solidFill(COLORS.totalBg);
      cell.font = FONT_GRAND;
      cell.alignment = CENTER;
      cell.border = MEDIUM_BORDER_BOTTOM;
    }
    col += 3;
  }
  ws.getCell(row, summaryStart).value = allQty;
  ws.getCell(row, summaryStart + 1).value = allRtn;
  ws.getCell(row, summaryStart + 2).value = allQty - allRtn;
  for (let j = 0; j < 3; j++) {
    const cell = ws.getCell(row, summaryStart + j);
    cell.numFmt = '#,##0';
    cell.fill = solidFill(COLORS.totalBg);
    cell.font = FONT_GRAND;
    cell.alignment = CENTER;
    cell.border = MEDIUM_BORDER_BOTTOM;
  }

  // ---- Column widths ----
  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 32;
  for (let c = zoneColStart; c <= lastCol; c++) {
    ws.getColumn(c).width = 10;
  }
}

// =================================================================
// Build per-zone sheet — products grouped by class with qty/rtn/net
// =================================================================
function buildZoneSheet(ws, zoneName, zoneProductData, productsByClass) {
  let row = 1;
  const COLS = 6; // 序號, 產品名稱, 類別, 訂量, 退量, 淨出

  // ---- Title ----
  ws.mergeCells(row, 1, row, 3);
  ws.getCell(row, 1).value = `${zoneName} — 其它類別各產品出貨統計`;
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

  // ---- Header ----
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

  // ---- Write products by section ----
  function writeSection(sectionName, products, sectionColor, sectionLightBg, classLabel) {
    // Section header
    ws.mergeCells(row, 1, row, COLS);
    ws.getCell(row, 1).value = `  ${sectionName}`;
    ws.getCell(row, 1).font = FONT_SECTION;
    ws.getCell(row, 1).fill = solidFill(sectionColor);
    ws.getCell(row, 1).alignment = LEFT_AL;
    for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
    row++;

    let sectionQty = 0, sectionRtn = 0;
    let rowIdx = 1;

    for (const p of products) {
      const d = zoneProductData[p] || { qty: 0, rtn: 0 };
      const net = d.qty - d.rtn;
      // Skip products with zero data in this zone
      if (d.qty === 0 && d.rtn === 0) continue;

      const bgColor = rowIdx % 2 !== 0 ? sectionLightBg : COLORS.dataBgEven;

      ws.getCell(row, 1).value = rowIdx;
      ws.getCell(row, 2).value = displayName(p);
      ws.getCell(row, 3).value = classLabel;
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

      sectionQty += d.qty;
      sectionRtn += d.rtn;
      rowIdx++;
      row++;
    }

    // Section subtotal
    const sectionNet = sectionQty - sectionRtn;

    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = `${sectionName} 小計`;
    ws.getCell(row, 3).value = '';
    ws.getCell(row, 4).value = sectionQty;
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = sectionRtn;
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).value = sectionNet;
    ws.getCell(row, 6).numFmt = '#,##0';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(sectionLightBg),
        font: FONT_TOTAL,
        border: MEDIUM_BORDER_BOTTOM,
        alignment: c <= 3 ? LEFT_AL : CENTER,
      });
    }
    row++;

    return { qty: sectionQty, rtn: sectionRtn };
  }

  let grandQty = 0, grandRtn = 0;

  for (let ci = 0; ci < PRODUCT_CLASSES.length; ci++) {
    const pc = PRODUCT_CLASSES[ci];
    const secColor = SECTION_COLORS[ci];
    const products = productsByClass[pc.key];

    const t = writeSection(pc.label, products, secColor.color, secColor.lightBg, pc.label);
    grandQty += t.qty;
    grandRtn += t.rtn;

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
}

// =================================================================
// Main
// =================================================================
async function main() {
  const zoneProductData = aggregateByZoneProduct(allRecords);
  const productsByClass = getProductsByClass(allRecords);

  const wb = new ExcelJS.Workbook();
  wb.creator = '金安出版社';
  wb.created = new Date();

  // Sheet 1: 總表
  const summaryWs = wb.addWorksheet('總表', {
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  buildSummarySheet(summaryWs, zoneProductData, productsByClass);

  // Sheets 2-4: Per zone
  for (const zone of TARGET_ZONES) {
    const ws = wb.addWorksheet(zone, {
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    buildZoneSheet(ws, zone, zoneProductData[zone] || {}, productsByClass);
  }

  const outputDir = path.join(__dirname, 'Output2');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = path.join(outputDir, '114上_其它類別_各區統計.xlsx');
  await wb.xlsx.writeFile(outputPath);
  console.log(`✓ ${outputPath}`);
  console.log(`  Sheets: ${wb.worksheets.map(ws => ws.name).join(', ')}`);

  // Print summary
  console.log('\n--- Product counts per class ---');
  for (const pc of PRODUCT_CLASSES) {
    console.log(`  ${pc.label}: ${productsByClass[pc.key].length} products`);
  }

  console.log('\n--- Zone totals ---');
  for (const zone of TARGET_ZONES) {
    const data = zoneProductData[zone] || {};
    let qty = 0, rtn = 0;
    for (const vals of Object.values(data)) {
      qty += vals.qty;
      rtn += vals.rtn;
    }
    console.log(`  ${zone}: qty=${qty}, rtn=${rtn}, net=${qty - rtn}`);
  }
  console.log('\nDone!');
}

main().catch(console.error);
