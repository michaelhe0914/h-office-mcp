// Generate styled XLSX: 114上 複習講義+新思維 — 北區、中區、南區
// Structure:
//   Sheet 1: 總表 (all zones summary)
//   Sheet 2: 北區 (per-product detail with qty/rtn/net)
//   Sheet 3: 中區
//   Sheet 4: 南區

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_all_regions_review.json'), 'utf-8'));

// =================================================================
// Constants
// =================================================================
const TARGET_ZONES = ['北區', '中區', '南區'];
const DATE_RANGE = '114.4.26~114.9.25';

// Canonical product order
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

const ALL_PRODUCTS = [...REVIEW_PRODUCTS, ...XINSWEI_PRODUCTS];

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
  // Section colors
  sxColor:      'FF70AD47',  // 雙向=綠
  sxLightBg:    'FFE2EFDA',
  q735Color:    'FFED7D31',  // 735=橘
  q735LightBg:  'FFFCE4D6',
  themeColor:   'FF5B9BD5',  // 主題讚=藍
  themeLightBg: 'FFDAEEF3',
  xinswColor:   'FF7030A0',  // 新思維=紫
  xinswLightBg: 'FFE8D5F5',
  // Summary
  totalBg:      'FFFFF2CC',
  totalFont:    'FF833C0B',
  grandTotalBg: 'FFDCE6F1',
  grandTotalFont:'FF1F4E79',
  // Data rows
  dataBgEven:   'FFF2F2F2',
  dataBgOdd:    'FFFFFFFF',
  borderColor:  'FF000000',
  // Zone colors for 總表
  northBg:      'FFD6E4F0',  // 淡藍
  centralBg:    'FFE2EFDA',  // 淡綠
  southBg:      'FFFCE4D6',  // 淡橘
  negativeBg:   'FFFFC7CE',  // 負數紅
};

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
const MEDIUM_BORDER_TOP = {
  top:    { style: 'medium', color: { argb: COLORS.borderColor } },
  left:   { style: 'thin',   color: { argb: COLORS.borderColor } },
  bottom: { style: 'thin',   color: { argb: COLORS.borderColor } },
  right:  { style: 'thin',   color: { argb: COLORS.borderColor } },
};

const CENTER = { horizontal: 'center', vertical: 'middle' };
const LEFT_AL = { horizontal: 'left', vertical: 'middle' };

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

function styleRange(ws, r1, c1, r2, c2, opts) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      styleCell(ws, r, c, opts);
    }
  }
}

// Short display name
function shortName(product) {
  return product
    .replace(/^\*/, '')
    .replace('735輕鬆讀 圖解複習', '735複習')
    .replace('(主題讚)複習', '主題讚複習');
}

// Product group for section coloring
function getGroup(product) {
  if (product.startsWith('*雙向')) return '雙向';
  if (product.startsWith('*735')) return '735';
  if (product.startsWith('*(主題讚)')) return '主題讚';
  if (product.startsWith('*新思維')) return '新思維';
  return '其他';
}

function getGroupColor(group) {
  switch (group) {
    case '雙向': return COLORS.sxColor;
    case '735': return COLORS.q735Color;
    case '主題讚': return COLORS.themeColor;
    case '新思維': return COLORS.xinswColor;
    default: return COLORS.headerBg;
  }
}

function getGroupLightBg(group) {
  switch (group) {
    case '雙向': return COLORS.sxLightBg;
    case '735': return COLORS.q735LightBg;
    case '主題讚': return COLORS.themeLightBg;
    case '新思維': return COLORS.xinswLightBg;
    default: return COLORS.dataBgOdd;
  }
}

// =================================================================
// Data aggregation
// =================================================================
function aggregateByZoneProduct(records) {
  // zone -> product -> { qty, rtn }
  const data = {};
  for (const r of records) {
    if (!TARGET_ZONES.includes(r.zone)) continue;
    if (!data[r.zone]) data[r.zone] = {};
    if (!data[r.zone][r.product]) data[r.zone][r.product] = { qty: 0, rtn: 0 };
    data[r.zone][r.product].qty += r.qty;
    data[r.zone][r.product].rtn += r.rtn_qty;
  }
  return data;
}

// =================================================================
// Build 總表 sheet
// =================================================================
function buildSummarySheet(ws, zoneData) {
  let row = 1;
  // Columns: A=序號, B=產品名稱, C=類別,
  // then for each zone: 訂量, 退量, 淨出 (3 cols each)
  // then: 全區訂量合計, 全區退量合計, 全區淨出合計
  const zoneColStart = 4; // col 4 starts zone data
  const colsPerZone = 3;
  const totalZoneCols = TARGET_ZONES.length * colsPerZone;
  const summaryStart = zoneColStart + totalZoneCols;
  const lastCol = summaryStart + 2; // 3 summary cols

  // ---- Title Row ----
  ws.mergeCells(row, 1, row, Math.floor(lastCol / 2));
  ws.getCell(row, 1).value = '114上 複習講義+新思維 各區出貨統計';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = CENTER;

  ws.mergeCells(row, Math.floor(lastCol / 2) + 1, row, lastCol);
  ws.getCell(row, Math.floor(lastCol / 2) + 1).value = DATE_RANGE;
  ws.getCell(row, Math.floor(lastCol / 2) + 1).font = FONT_DATE;
  ws.getCell(row, Math.floor(lastCol / 2) + 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, Math.floor(lastCol / 2) + 1).alignment = { horizontal: 'right', vertical: 'middle' };

  for (let c = 1; c <= lastCol; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) cell.fill = solidFill(COLORS.titleBg);
    cell.border = THIN_BORDER;
  }
  row++;

  // ---- Zone Header Row (row 2) ----
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';

  const zoneBgs = [COLORS.northBg, COLORS.centralBg, COLORS.southBg];
  let col = zoneColStart;
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    ws.mergeCells(row, col, row, col + 2);
    ws.getCell(row, col).value = TARGET_ZONES[i];
    ws.getCell(row, col).font = FONT_SUBHDR;
    ws.getCell(row, col).fill = solidFill(zoneBgs[i]);
    ws.getCell(row, col).alignment = CENTER;
    col += 3;
  }

  // Total zone header
  ws.mergeCells(row, summaryStart, row, summaryStart + 2);
  ws.getCell(row, summaryStart).value = '全區合計';
  ws.getCell(row, summaryStart).font = FONT_SUBHDR;
  ws.getCell(row, summaryStart).fill = solidFill(COLORS.grandTotalBg);
  ws.getCell(row, summaryStart).alignment = CENTER;

  for (let c = 1; c <= lastCol; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.border) cell.border = THIN_BORDER;
    if (!cell.fill || !cell.fill.fgColor) {
      cell.fill = solidFill(COLORS.subHeaderBg);
    }
    cell.font = cell.font || FONT_SUBHDR;
    cell.border = THIN_BORDER;
  }
  row++;

  // ---- Sub-header Row (row 3): column labels ----
  ws.getCell(row, 1).value = '序號';
  ws.getCell(row, 2).value = '產品名稱';
  ws.getCell(row, 3).value = '類別';

  col = zoneColStart;
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    ws.getCell(row, col).value = '訂量';
    ws.getCell(row, col + 1).value = '退量';
    ws.getCell(row, col + 2).value = '淨出';
    for (let j = 0; j < 3; j++) {
      ws.getCell(row, col + j).fill = solidFill(zoneBgs[i]);
    }
    col += 3;
  }
  ws.getCell(row, summaryStart).value = '訂量';
  ws.getCell(row, summaryStart + 1).value = '退量';
  ws.getCell(row, summaryStart + 2).value = '淨出';

  for (let c = 1; c <= lastCol; c++) {
    styleCell(ws, row, c, {
      fill: ws.getCell(row, c).fill?.fgColor ? undefined : solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER,
    });
  }
  // Override fill for zone sub-headers to keep zone color but darker
  col = zoneColStart;
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    for (let j = 0; j < 3; j++) {
      ws.getCell(row, col + j).fill = solidFill(zoneBgs[i]);
      ws.getCell(row, col + j).font = FONT_SUBHDR;
    }
    col += 3;
  }
  for (let j = 0; j < 3; j++) {
    ws.getCell(row, summaryStart + j).fill = solidFill(COLORS.grandTotalBg);
    ws.getCell(row, summaryStart + j).font = FONT_SUBHDR;
  }
  row++;

  // ---- Data Rows ----
  let prevGroup = '';
  let idx = 1;
  const grandTotals = TARGET_ZONES.map(() => ({ qty: 0, rtn: 0 }));
  const allZoneTotals = { qty: 0, rtn: 0 };

  // Subtotals by category
  const categoryTotals = {};
  
  for (let pi = 0; pi < ALL_PRODUCTS.length; pi++) {
    const p = ALL_PRODUCTS[pi];
    const group = getGroup(p);
    const nextGroup = pi < ALL_PRODUCTS.length - 1 ? getGroup(ALL_PRODUCTS[pi + 1]) : null;

    // Track category totals
    if (!categoryTotals[group]) {
      categoryTotals[group] = TARGET_ZONES.map(() => ({ qty: 0, rtn: 0 }));
    }

    const bgColor = idx % 2 === 0 ? COLORS.dataBgEven : COLORS.dataBgOdd;
    
    ws.getCell(row, 1).value = idx;
    ws.getCell(row, 2).value = shortName(p);
    ws.getCell(row, 3).value = group;

    col = zoneColStart;
    let rowTotalQty = 0, rowTotalRtn = 0;
    for (let i = 0; i < TARGET_ZONES.length; i++) {
      const zone = TARGET_ZONES[i];
      const d = zoneData[zone]?.[p] || { qty: 0, rtn: 0 };
      ws.getCell(row, col).value = d.qty || '';
      ws.getCell(row, col + 1).value = d.rtn || '';
      ws.getCell(row, col + 2).value = d.qty - d.rtn || '';
      if (d.qty) ws.getCell(row, col).numFmt = '#,##0';
      if (d.rtn) ws.getCell(row, col + 1).numFmt = '#,##0';
      if (d.qty - d.rtn) ws.getCell(row, col + 2).numFmt = '#,##0';
      
      grandTotals[i].qty += d.qty;
      grandTotals[i].rtn += d.rtn;
      categoryTotals[group][i].qty += d.qty;
      categoryTotals[group][i].rtn += d.rtn;
      rowTotalQty += d.qty;
      rowTotalRtn += d.rtn;
      col += 3;
    }

    // All zones total
    ws.getCell(row, summaryStart).value = rowTotalQty || '';
    ws.getCell(row, summaryStart + 1).value = rowTotalRtn || '';
    ws.getCell(row, summaryStart + 2).value = rowTotalQty - rowTotalRtn || '';
    if (rowTotalQty) ws.getCell(row, summaryStart).numFmt = '#,##0';
    if (rowTotalRtn) ws.getCell(row, summaryStart + 1).numFmt = '#,##0';
    if (rowTotalQty - rowTotalRtn) ws.getCell(row, summaryStart + 2).numFmt = '#,##0';

    allZoneTotals.qty += rowTotalQty;
    allZoneTotals.rtn += rowTotalRtn;

    // Styling
    for (let c = 1; c <= lastCol; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER,
        alignment: c <= 3 ? LEFT_AL : CENTER,
      });
    }
    // Color group label cell
    ws.getCell(row, 3).font = { ...FONT_DATA, bold: true, color: { argb: getGroupColor(group).replace('FF', 'FF') } };

    // Summary columns slightly highlighted
    for (let j = 0; j < 3; j++) {
      ws.getCell(row, summaryStart + j).fill = solidFill(idx % 2 === 0 ? 'FFEDF2F9' : 'FFF5F8FC');
      ws.getCell(row, summaryStart + j).font = FONT_BOLD;
    }

    idx++;
    row++;

    // Insert category subtotal row when group changes
    if (nextGroup !== group) {
      const catTotals = categoryTotals[group];
      ws.getCell(row, 1).value = '';
      ws.getCell(row, 2).value = `${group} 小計`;
      ws.getCell(row, 3).value = '';
      
      col = zoneColStart;
      let catAllQty = 0, catAllRtn = 0;
      for (let i = 0; i < TARGET_ZONES.length; i++) {
        ws.getCell(row, col).value = catTotals[i].qty;
        ws.getCell(row, col).numFmt = '#,##0';
        ws.getCell(row, col + 1).value = catTotals[i].rtn;
        ws.getCell(row, col + 1).numFmt = '#,##0';
        ws.getCell(row, col + 2).value = catTotals[i].qty - catTotals[i].rtn;
        ws.getCell(row, col + 2).numFmt = '#,##0';
        catAllQty += catTotals[i].qty;
        catAllRtn += catTotals[i].rtn;
        col += 3;
      }
      ws.getCell(row, summaryStart).value = catAllQty;
      ws.getCell(row, summaryStart).numFmt = '#,##0';
      ws.getCell(row, summaryStart + 1).value = catAllRtn;
      ws.getCell(row, summaryStart + 1).numFmt = '#,##0';
      ws.getCell(row, summaryStart + 2).value = catAllQty - catAllRtn;
      ws.getCell(row, summaryStart + 2).numFmt = '#,##0';

      const subBg = getGroupLightBg(group);
      for (let c = 1; c <= lastCol; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(subBg),
          font: FONT_TOTAL,
          border: MEDIUM_BORDER_BOTTOM,
          alignment: c <= 3 ? LEFT_AL : CENTER,
        });
      }
      ws.getCell(row, 2).font = { ...FONT_TOTAL, color: { argb: getGroupColor(group).replace('FF', 'FF') } };
      row++;
    }

    prevGroup = group;
  }

  // ---- Grand Total Row ----
  row++; // blank separator
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '總合計';
  ws.getCell(row, 3).value = '';
  
  col = zoneColStart;
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    ws.getCell(row, col).value = grandTotals[i].qty;
    ws.getCell(row, col).numFmt = '#,##0';
    ws.getCell(row, col + 1).value = grandTotals[i].rtn;
    ws.getCell(row, col + 1).numFmt = '#,##0';
    ws.getCell(row, col + 2).value = grandTotals[i].qty - grandTotals[i].rtn;
    ws.getCell(row, col + 2).numFmt = '#,##0';
    col += 3;
  }
  ws.getCell(row, summaryStart).value = allZoneTotals.qty;
  ws.getCell(row, summaryStart).numFmt = '#,##0';
  ws.getCell(row, summaryStart + 1).value = allZoneTotals.rtn;
  ws.getCell(row, summaryStart + 1).numFmt = '#,##0';
  ws.getCell(row, summaryStart + 2).value = allZoneTotals.qty - allZoneTotals.rtn;
  ws.getCell(row, summaryStart + 2).numFmt = '#,##0';

  for (let c = 1; c <= lastCol; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c <= 3 ? CENTER : CENTER,
    });
  }

  // ---- Column widths ----
  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 26;
  ws.getColumn(3).width = 8;
  for (let c = zoneColStart; c <= lastCol; c++) {
    ws.getColumn(c).width = 10;
  }
  ws.getRow(1).height = 28;
}

// =================================================================
// Build per-zone sheet
// =================================================================
function buildZoneSheet(ws, zoneName, zoneProductData, zoneBg) {
  let row = 1;
  const COLS = 6; // 產品名稱, 類別, 訂量, 退量, 淨出, 退書率

  // ---- Title ----
  ws.mergeCells(row, 1, row, 3);
  ws.getCell(row, 1).value = `${zoneName} — 複習講義+新思維 出貨統計`;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT_AL;

  ws.mergeCells(row, 4, row, COLS);
  ws.getCell(row, 4).value = DATE_RANGE;
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

  // ---- Write products by section ----
  function writeSection(sectionName, products, sectionColor, sectionLightBg) {
    // Section header
    ws.mergeCells(row, 1, row, COLS);
    ws.getCell(row, 1).value = `  ${sectionName}`;
    ws.getCell(row, 1).font = FONT_SECTION;
    ws.getCell(row, 1).fill = solidFill(sectionColor);
    ws.getCell(row, 1).alignment = LEFT_AL;
    for (let c = 1; c <= COLS; c++) {
      ws.getCell(row, c).border = THIN_BORDER;
    }
    row++;

    let sectionQty = 0, sectionRtn = 0;
    let rowIdx = 0;

    for (const p of products) {
      const d = zoneProductData[p] || { qty: 0, rtn: 0 };
      const net = d.qty - d.rtn;
      const rtnRate = d.qty > 0 ? d.rtn / d.qty : 0;
      const bgColor = rowIdx % 2 === 0 ? sectionLightBg : COLORS.dataBgOdd;

      ws.getCell(row, 1).value = shortName(p);
      ws.getCell(row, 2).value = getGroup(p);
      ws.getCell(row, 3).value = d.qty || '';
      if (d.qty) ws.getCell(row, 3).numFmt = '#,##0';
      ws.getCell(row, 4).value = d.rtn || '';
      if (d.rtn) ws.getCell(row, 4).numFmt = '#,##0';
      ws.getCell(row, 5).value = net;
      ws.getCell(row, 5).numFmt = '#,##0';
      ws.getCell(row, 6).value = rtnRate > 0 ? rtnRate : '';
      if (rtnRate > 0) ws.getCell(row, 6).numFmt = '0.0%';

      for (let c = 1; c <= COLS; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(bgColor),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c <= 2 ? LEFT_AL : CENTER,
        });
      }

      if (net < 0) {
        ws.getCell(row, 5).fill = solidFill(COLORS.negativeBg);
        ws.getCell(row, 5).font = { ...FONT_DATA, color: { argb: 'FFCC0000' } };
      }

      sectionQty += d.qty;
      sectionRtn += d.rtn;
      rowIdx++;
      row++;
    }

    // Section subtotal
    const sectionNet = sectionQty - sectionRtn;
    const sectionRate = sectionQty > 0 ? sectionRtn / sectionQty : 0;

    ws.getCell(row, 1).value = `${sectionName} 小計`;
    ws.getCell(row, 2).value = '';
    ws.getCell(row, 3).value = sectionQty;
    ws.getCell(row, 3).numFmt = '#,##0';
    ws.getCell(row, 4).value = sectionRtn;
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = sectionNet;
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).value = sectionRate > 0 ? sectionRate : '';
    if (sectionRate > 0) ws.getCell(row, 6).numFmt = '0.0%';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(sectionLightBg),
        font: FONT_TOTAL,
        border: MEDIUM_BORDER_BOTTOM,
        alignment: c <= 2 ? LEFT_AL : CENTER,
      });
    }
    row++;

    return { qty: sectionQty, rtn: sectionRtn };
  }

  // Group products by section
  const sxProducts = REVIEW_PRODUCTS.filter(p => getGroup(p) === '雙向');
  const q735Products = REVIEW_PRODUCTS.filter(p => getGroup(p) === '735');
  const themeProducts = REVIEW_PRODUCTS.filter(p => getGroup(p) === '主題讚');

  const t1 = writeSection('雙向複習', sxProducts, COLORS.sxColor, COLORS.sxLightBg);
  row++; // blank
  const t2 = writeSection('735複習', q735Products, COLORS.q735Color, COLORS.q735LightBg);
  row++;
  const t3 = writeSection('主題讚複習', themeProducts, COLORS.themeColor, COLORS.themeLightBg);
  
  // 複習講義 total
  row++;
  const reviewQty = t1.qty + t2.qty + t3.qty;
  const reviewRtn = t1.rtn + t2.rtn + t3.rtn;
  ws.getCell(row, 1).value = '複習講義 合計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = reviewQty;
  ws.getCell(row, 3).numFmt = '#,##0';
  ws.getCell(row, 4).value = reviewRtn;
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = reviewQty - reviewRtn;
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = reviewQty > 0 ? reviewRtn / reviewQty : '';
  if (reviewQty > 0) ws.getCell(row, 6).numFmt = '0.0%';
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.grandTotalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c <= 2 ? CENTER : CENTER,
    });
  }
  row += 2;

  const t4 = writeSection('新思維複習', XINSWEI_PRODUCTS, COLORS.xinswColor, COLORS.xinswLightBg);

  // ---- Grand Total ----
  row++;
  const grandQty = reviewQty + t4.qty;
  const grandRtn = reviewRtn + t4.rtn;

  ws.getCell(row, 1).value = '總合計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = grandQty;
  ws.getCell(row, 3).numFmt = '#,##0';
  ws.getCell(row, 4).value = grandRtn;
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = grandQty - grandRtn;
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = grandQty > 0 ? grandRtn / grandQty : '';
  if (grandQty > 0) ws.getCell(row, 6).numFmt = '0.0%';
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
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 12;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 10;
}

// =================================================================
// Main
// =================================================================
async function main() {
  const zoneData = aggregateByZoneProduct(allRecords);
  
  const wb = new ExcelJS.Workbook();
  wb.creator = '金安出版社';
  wb.created = new Date();

  // Sheet 1: 總表
  const summaryWs = wb.addWorksheet('總表', {
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  buildSummarySheet(summaryWs, zoneData);

  // Sheets 2-4: Per zone
  const zoneBgs = [COLORS.northBg, COLORS.centralBg, COLORS.southBg];
  for (let i = 0; i < TARGET_ZONES.length; i++) {
    const zone = TARGET_ZONES[i];
    const ws = wb.addWorksheet(zone, {
      pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    buildZoneSheet(ws, zone, zoneData[zone] || {}, zoneBgs[i]);
  }

  const outputDir = path.join(__dirname, 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = path.join(outputDir, '114上_複習講義_新思維_各區統計.xlsx');
  await wb.xlsx.writeFile(outputPath);
  console.log(`✓ ${outputPath}`);
  console.log(`  Sheets: ${wb.worksheets.map(ws => ws.name).join(', ')}`);
  console.log('Done!');
}

main().catch(console.error);
