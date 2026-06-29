// Generate styled xlsx files for 國中考卷 (exam papers)
// Each sales rep gets one xlsx with customer sheets
// Plus a summary xlsx combining all customers across all sales reps

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_north_exam.json'), 'utf-8'));
const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉瑋", "朱鵬學"];

// =================================================================
// Color Palette & Style Definitions
// =================================================================
const COLORS = {
  titleBg:     'FF1F4E79',  // 深藍
  titleFont:   'FFFFFFFF',
  headerBg:    'FF4472C4',  // 中藍
  headerFont:  'FFFFFFFF',
  subHeaderBg: 'FFD6E4F0',  // 淺藍 (grade row)
  subHeaderFont:'FF1F4E79',
  totalBg:     'FFFFF2CC',  // 合計=淡黃
  totalFont:   'FF833C0B',
  borderColor: 'FF000000',
  summaryColBg:'FFDCE6F1',  // 訂書/退貨/實際出貨欄
  // Version backgrounds
  kVersionBg:  'FFDCE6D0',  // K版淡綠
  nVersionBg:  'FFFFFFFF',  // N版白色
  hVersionBg:  'FFFFFFFF',  // H版白色
  // Section colors for K/N/H sections
  kSectionColor: 'FF548235',  // K版=深綠
  nSectionColor: 'FFFFFFFF',  // N版白色
  hSectionColor: 'FFFFFFFF',  // H版白色
  whiteExamBg:   'FFF2F2F2',  // 白卷=淺灰
  whiteExamColor:'FF808080',  // 白卷=灰
};

const FONT_TITLE = { name: '微軟正黑體', size: 14, bold: true, color: { argb: COLORS.titleFont } };
const FONT_HEADER = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.headerFont } };
const FONT_SUBHEADER = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.subHeaderFont } };
const FONT_DATA = { name: '微軟正黑體', size: 10 };
const FONT_TOTAL = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.totalFont } };
const FONT_SECTION_LABEL = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };

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

// =================================================================
// Product parsing for 考卷
// =================================================================
function parseExamProduct(product, productClass) {
  // Extract volume (冊數) from product name
  const volMatch = product.match(/[（(]([１-６1-6])[）)]/);
  if (!volMatch) return null;
  const volMap = { '１': 1, '２': 2, '３': 3, '４': 4, '５': 5, '６': 6 };
  const volume = volMap[volMatch[1]] || parseInt(volMatch[1]);

  // Determine version (K/N/H) and variant (a/b/合計) from productClass
  let version, variant;
  const isWhiteExam = productClass.includes('白卷');

  if (isWhiteExam) {
    // 白卷 product class mixes all publishers - need to parse from product name
    if (product.includes('康版') || product.includes('[康版]')) {
      version = 'K';
      if (product.includes('A卷')) variant = 'a';
      else if (product.includes('B卷')) variant = 'b';
      else variant = '合計';
    } else if (product.includes('南版')) {
      version = 'N';
      variant = '合計';
    } else if (product.includes('翰版')) {
      version = 'H';
      if (product.includes('A卷')) variant = 'a';
      else if (product.includes('B卷')) variant = 'b';
      else variant = '合計';
    } else {
      return null;
    }
  } else if (productClass.includes('康卷')) {
    version = 'K';
    variant = productClass.includes('A卷') ? 'a' : 'b';
  } else if (productClass.includes('南卷')) {
    version = 'N';
    variant = '合計';
  } else if (productClass.includes('翰卷')) {
    version = 'H';
    variant = productClass.includes('A卷') ? 'a' : 'b';
  } else {
    return null;
  }

  // Extract subject from product name
  let subject;
  if (product.includes('國文')) subject = '國文';
  else if (product.includes('英語') || product.includes('英文')) subject = '英文';
  else if (product.includes('數學')) subject = '數學';
  else if (product.includes('自然')) subject = '自然';
  else if (product.includes('地理')) subject = '地理';
  else if (product.includes('歷史')) subject = '歷史';
  else if (product.includes('公民')) subject = '公民';
  else return null;

  // Volume to grade mapping
  let grade;
  if (volume <= 2) grade = '一';
  else if (volume <= 4) grade = '三';
  else grade = '五';

  return { version, variant, subject, volume, grade, isWhiteExam };
}

// Build data structure:
// salesData[sales][customer][version][variant][subject][grade] = qty
// Also track white exam separately: whiteExamData[sales][customer][version][variant][subject][grade] = qty
function buildExamData(records) {
  const salesData = {};
  for (const r of records) {
    const parsed = parseExamProduct(r.product, r.productClass);
    if (!parsed) {
      // Log unrecognized products for debugging
      // console.warn(`  Unrecognized: ${r.product} [${r.productClass}]`);
      continue;
    }
    const { version, variant, subject, grade, isWhiteExam } = parsed;
    const dataType = isWhiteExam ? '白卷' : '考卷';

    if (!salesData[r.sales]) salesData[r.sales] = {};
    if (!salesData[r.sales][r.customer]) salesData[r.sales][r.customer] = {};
    if (!salesData[r.sales][r.customer][dataType]) salesData[r.sales][r.customer][dataType] = {};
    if (!salesData[r.sales][r.customer][dataType][version]) salesData[r.sales][r.customer][dataType][version] = {};
    if (!salesData[r.sales][r.customer][dataType][version][variant]) salesData[r.sales][r.customer][dataType][version][variant] = {};
    if (!salesData[r.sales][r.customer][dataType][version][variant][subject]) salesData[r.sales][r.customer][dataType][version][variant][subject] = {};
    salesData[r.sales][r.customer][dataType][version][variant][subject][grade] =
      (salesData[r.sales][r.customer][dataType][version][variant][subject][grade] || 0) + r.qty;
  }
  return salesData;
}

// =================================================================
// Style helpers
// =================================================================
function styleCell(ws, row, col, { fill, font, border, alignment, numFmt } = {}) {
  const cell = ws.getCell(row, col);
  if (fill) cell.fill = fill;
  if (font) cell.font = font;
  if (border) cell.border = border;
  if (alignment) cell.alignment = alignment;
  if (numFmt) cell.numFmt = numFmt;
}

function solidFill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function colLetter(colNum) {
  let temp = colNum;
  let letter = '';
  while (temp > 0) {
    let modulo = (temp - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    temp = Math.floor((temp - modulo) / 26);
  }
  return letter;
}

// =================================================================
// Build a styled sheet for one customer
// =================================================================
const SUBJECTS = ['國文', '英文', '數學', '自然', '地理', '歷史', '公民'];
const GRADES = ['一', '三', '五'];
const NUM_SUBJ_COLS = SUBJECTS.length * 3; // 21
const TOTAL_COLS = 3 + NUM_SUBJ_COLS + 3; // 品項+卷別+科目 + 21 data + 訂書+退貨+實際出貨 = 27

async function buildStyledSheet(ws, customer, data) {
  let row = 1;
  const examData = data['考卷'] || {};
  const whiteData = data['白卷'] || {};

  // ==========================================
  // Title Row
  // ==========================================
  ws.mergeCells(row, 1, row, 6);
  ws.getCell(row, 1).value = '114年上期\r\n 單冊考卷';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = { ...CENTER, wrapText: true };

  ws.mergeCells(row, 7, row, 12);
  ws.getCell(row, 7).value = customer;
  ws.getCell(row, 7).font = { ...FONT_TITLE, size: 12 };
  ws.getCell(row, 7).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 7).alignment = CENTER;

  ws.mergeCells(row, 13, row, TOTAL_COLS);
  ws.getCell(row, 13).value = '114.4.26~114.9.25';
  ws.getCell(row, 13).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 13).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 13).alignment = { horizontal: 'right', vertical: 'middle' };

  for (let c = 1; c <= TOTAL_COLS; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) cell.fill = solidFill(COLORS.titleBg);
    cell.border = THIN_BORDER;
  }
  row++;

  // ==========================================
  // Header Row 1: Subject groups
  // ==========================================
  ws.getCell(row, 1).value = '品項';
  ws.getCell(row, 2).value = '卷別';
  ws.getCell(row, 3).value = '科目';
  let col = 4;
  for (const subj of SUBJECTS) {
    ws.mergeCells(row, col, row, col + 2);
    ws.getCell(row, col).value = subj;
    ws.getCell(row, col).alignment = CENTER;
    col += 3;
  }
  ws.getCell(row, col).value = '訂書';
  ws.getCell(row, col + 1).value = '退\r\n貨';
  ws.getCell(row, col + 2).value = '實際';

  // Style header row
  for (let c = 1; c <= TOTAL_COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER,
    });
  }
  // Merge fixed-cols with next row
  ws.mergeCells(row, 1, row + 1, 1);
  ws.mergeCells(row, 2, row + 1, 2);
  ws.mergeCells(row, 3, row + 1, 3);
  // Merge summary cols with next row
  ws.mergeCells(row, col, row + 1, col);
  ws.mergeCells(row, col + 1, row + 1, col + 1);
  ws.mergeCells(row, col + 2, row + 1, col + 2);
  row++;

  // ==========================================
  // Header Row 2: Grade labels
  // ==========================================
  col = 4;
  for (const subj of SUBJECTS) {
    for (const grade of GRADES) {
      ws.getCell(row, col).value = grade;
      col++;
    }
  }
  for (let c = 1; c <= TOTAL_COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.subHeaderBg),
      font: FONT_SUBHEADER,
      border: THIN_BORDER,
      alignment: CENTER,
    });
  }
  row++;

  // ==========================================
  // Helper: write one data row
  // ==========================================
  function writeDataRow(versionLabel, variantLabel, areaLabel, rowData, bgColor) {
    ws.getCell(row, 1).value = versionLabel;
    ws.getCell(row, 2).value = variantLabel;
    ws.getCell(row, 3).value = areaLabel;

    let rowTotal = 0;
    let col = 4;
    for (const subj of SUBJECTS) {
      for (const grade of GRADES) {
        const qty = (rowData[subj] && rowData[subj][grade]) || 0;
        ws.getCell(row, col).value = qty || '';
        if (qty > 0) ws.getCell(row, col).numFmt = '#,##0';
        rowTotal += qty;
        col++;
      }
    }

    // Summary columns
    ws.getCell(row, col).value = rowTotal || 0;
    ws.getCell(row, col).numFmt = '#,##0';
    ws.getCell(row, col + 1).value = '';
    ws.getCell(row, col + 2).value = rowTotal || 0;
    ws.getCell(row, col + 2).numFmt = '#,##0';

    // Style entire row
    for (let c = 1; c <= TOTAL_COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER,
        alignment: c <= 3 ? CENTER : CENTER,
      });
    }
    // Summary columns highlight
    for (let sc = col; sc <= col + 2; sc++) {
      ws.getCell(row, sc).fill = solidFill(COLORS.summaryColBg);
      ws.getCell(row, sc).font = { ...FONT_DATA, bold: true };
    }

    const currentRow = row;
    row++;
    return currentRow;
  }

  // ==========================================
  // Helper: write total row (summing data across variants)
  // ==========================================
  function writeTotalRow(label, areaLabel, dataSets, bgColor, fontStyle, borderStyle) {
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = label;
    ws.getCell(row, 3).value = areaLabel;

    let grandTotal = 0;
    let col = 4;
    for (const subj of SUBJECTS) {
      for (const grade of GRADES) {
        let total = 0;
        for (const ds of dataSets) {
          total += (ds[subj] && ds[subj][grade]) || 0;
        }
        ws.getCell(row, col).value = total || 0;
        ws.getCell(row, col).numFmt = '#,##0';
        grandTotal += total;
        col++;
      }
    }

    ws.getCell(row, col).value = grandTotal || 0;
    ws.getCell(row, col).numFmt = '#,##0';
    ws.getCell(row, col + 1).value = '';
    ws.getCell(row, col + 2).value = grandTotal || 0;
    ws.getCell(row, col + 2).numFmt = '#,##0';

    for (let c = 1; c <= TOTAL_COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: fontStyle,
        border: borderStyle,
        alignment: CENTER,
      });
    }
    // Summary columns highlight
    for (let sc = col; sc <= col + 2; sc++) {
      ws.getCell(row, sc).fill = solidFill(COLORS.summaryColBg);
      ws.getCell(row, sc).font = { ...fontStyle, bold: true };
    }

    const currentRow = row;
    row++;
    return currentRow;
  }

  // ==========================================
  // K版 section (a卷, b卷, 總數)
  // ==========================================
  const kData = examData['K'] || {};
  const kWhite = whiteData['K'] || {};
  const kAData = kData['a'] || {};
  const kBData = kData['b'] || {};
  const kStartRow = row;

  // K版 label cell
  writeDataRow('K版', 'a卷', customer, kAData, COLORS.kVersionBg);
  writeDataRow('', 'b卷', customer, kBData, COLORS.kVersionBg);

  // K版 total row (sum of a + b)
  const allKDataSets = [kAData, kBData];
  // Also add white exam K data
  const kWhiteA = kWhite['a'] || {};
  const kWhiteB = kWhite['b'] || {};
  const allKWithWhite = [kAData, kBData, kWhiteA, kWhiteB];
  const kTotalRow = writeTotalRow('總數', '總數', allKWithWhite, COLORS.totalBg, FONT_TOTAL, THIN_BORDER);

  // Merge K版 label vertically
  ws.mergeCells(kStartRow, 1, kTotalRow, 1);
  ws.getCell(kStartRow, 1).font = FONT_SECTION_LABEL;
  ws.getCell(kStartRow, 1).fill = solidFill(COLORS.kSectionColor);
  ws.getCell(kStartRow, 1).alignment = CENTER;

  // ==========================================
  // N版 section (合計, single row)
  // ==========================================
  const nData = examData['N'] || {};
  const nWhite = whiteData['N'] || {};
  const nAllData = nData['合計'] || {};
  const nWhiteAll = nWhite['合計'] || {};
  // Merge N data + white N data
  const nMergedData = {};
  for (const subj of SUBJECTS) {
    nMergedData[subj] = {};
    for (const grade of GRADES) {
      nMergedData[subj][grade] = ((nAllData[subj] && nAllData[subj][grade]) || 0)
                                + ((nWhiteAll[subj] && nWhiteAll[subj][grade]) || 0);
    }
  }
  const nRow = writeDataRow('N版', '合計', customer, nMergedData, COLORS.nVersionBg);
  ws.getCell(nRow, 1).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FF000000' } };
  ws.getCell(nRow, 1).fill = solidFill(COLORS.nSectionColor);
  ws.getCell(nRow, 1).alignment = CENTER;

  // ==========================================
  // H版 section (a卷, b卷, 總數)
  // ==========================================
  const hData = examData['H'] || {};
  const hWhite = whiteData['H'] || {};
  const hAData = hData['a'] || {};
  const hBData = hData['b'] || {};
  const hStartRow = row;

  writeDataRow('H版', 'a卷', customer, hAData, COLORS.hVersionBg);
  writeDataRow('', 'b卷', customer, hBData, COLORS.hVersionBg);

  const hWhiteA = hWhite['a'] || {};
  const hWhiteB = hWhite['b'] || {};
  const allHWithWhite = [hAData, hBData, hWhiteA, hWhiteB];
  const hTotalRow = writeTotalRow('總數', '總數', allHWithWhite, COLORS.totalBg, FONT_TOTAL, THIN_BORDER);

  // Merge H版 label vertically
  ws.mergeCells(hStartRow, 1, hTotalRow, 1);
  ws.getCell(hStartRow, 1).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FF000000' } };
  ws.getCell(hStartRow, 1).fill = solidFill(COLORS.hSectionColor);
  ws.getCell(hStartRow, 1).alignment = CENTER;

  // ==========================================
  // Grand total row (all versions)
  // ==========================================
  const allDataSets = [];
  // Collect all data sets from all versions
  for (const ver of ['K', 'N', 'H']) {
    const vExam = examData[ver] || {};
    const vWhite = whiteData[ver] || {};
    for (const variant of Object.keys(vExam)) {
      allDataSets.push(vExam[variant]);
    }
    for (const variant of Object.keys(vWhite)) {
      allDataSets.push(vWhite[variant]);
    }
  }
  const grandTotalRow = writeTotalRow('', '總數', allDataSets, COLORS.totalBg, FONT_TOTAL, MEDIUM_BORDER_BOTTOM);

  // ==========================================
  // Summary table (bottom)
  // ==========================================
  row += 3; // spacing
  const summaryStartRow = row;

  // Title
  ws.mergeCells(row, 1, row + 1, 2);
  ws.getCell(row, 1).value = '總合計';
  ws.getCell(row, 1).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = CENTER;

  // Headers
  ws.getCell(row, 3).value = '冊數';
  ws.getCell(row, 3).font = FONT_HEADER;
  ws.getCell(row, 3).fill = solidFill(COLORS.headerBg);
  ws.getCell(row, 3).alignment = CENTER;

  const summaryHeaders = ['一', '三', '五', '總計'];
  for (let i = 0; i < summaryHeaders.length; i++) {
    ws.getCell(row, 4 + i).value = summaryHeaders[i];
    ws.getCell(row, 4 + i).font = FONT_HEADER;
    ws.getCell(row, 4 + i).fill = solidFill(COLORS.headerBg);
    ws.getCell(row, 4 + i).alignment = CENTER;
  }
  for (let c = 1; c <= 7; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
  }
  row++;

  // Data row: 總數
  ws.getCell(row, 3).value = '總數';
  ws.getCell(row, 3).font = { ...FONT_DATA, bold: true };
  ws.getCell(row, 3).fill = solidFill(COLORS.totalBg);
  ws.getCell(row, 3).alignment = CENTER;

  // Calculate totals per grade
  const gradeTotals = { '一': 0, '三': 0, '五': 0 };
  for (const ds of allDataSets) {
    for (const subj of SUBJECTS) {
      for (const grade of GRADES) {
        gradeTotals[grade] += (ds[subj] && ds[subj][grade]) || 0;
      }
    }
  }

  // Use formulas referencing the grand total row
  for (let gi = 0; gi < GRADES.length; gi++) {
    const g = GRADES[gi];
    // Sum all subject columns for this grade from the grand total row
    const formulaCols = [];
    for (let si = 0; si < SUBJECTS.length; si++) {
      const dataCol = 4 + si * 3 + gi;
      formulaCols.push(`${colLetter(dataCol)}${grandTotalRow}`);
    }
    ws.getCell(row, 4 + gi).value = { formula: `SUM(${formulaCols.join(',')})` };
    ws.getCell(row, 4 + gi).numFmt = '#,##0';
    ws.getCell(row, 4 + gi).font = FONT_DATA;
    ws.getCell(row, 4 + gi).fill = solidFill(COLORS.totalBg);
    ws.getCell(row, 4 + gi).alignment = CENTER;
  }

  // Total column: sum of all grades
  ws.getCell(row, 7).value = { formula: `SUM(D${row}:F${row})` };
  ws.getCell(row, 7).numFmt = '#,##0';
  ws.getCell(row, 7).font = FONT_TOTAL;
  ws.getCell(row, 7).fill = solidFill(COLORS.totalBg);
  ws.getCell(row, 7).alignment = CENTER;

  for (let c = 1; c <= 7; c++) {
    ws.getCell(row, c).border = MEDIUM_BORDER_BOTTOM;
  }
  row++;

  // ==========================================
  // Final pass: apply borders to ALL cells
  // ==========================================
  const lastRow = row;
  for (let r = 1; r <= lastRow; r++) {
    const limitCol = (r >= summaryStartRow) ? 7 : TOTAL_COLS;
    for (let c = 1; c <= limitCol; c++) {
      const cell = ws.getCell(r, c);
      if (!cell.border || !cell.border.top) {
        cell.border = THIN_BORDER;
      }
    }
  }

  // ==========================================
  // Column widths
  // ==========================================
  ws.getColumn(1).width = 8;
  ws.getColumn(2).width = 8;
  ws.getColumn(3).width = 10;
  for (let c = 4; c <= TOTAL_COLS; c++) {
    ws.getColumn(c).width = 7;
  }
  // Wider for summary cols
  ws.getColumn(TOTAL_COLS - 2).width = 8;  // 訂書
  ws.getColumn(TOTAL_COLS - 1).width = 6;  // 退貨
  ws.getColumn(TOTAL_COLS).width = 8;      // 實際出貨

  ws.getRow(1).height = 32;

  // Print setup
  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
  };
}

// =================================================================
// Aggregate customer data: merge all customers' data into one
// =================================================================
function aggregateCustomerData(customerDataMap) {
  const agg = { '考卷': {}, '白卷': {} };
  for (const customer of Object.keys(customerDataMap)) {
    for (const dataType of ['考卷', '白卷']) {
      const dt = customerDataMap[customer][dataType];
      if (!dt) continue;
      for (const ver of Object.keys(dt)) {
        if (!agg[dataType][ver]) agg[dataType][ver] = {};
        for (const variant of Object.keys(dt[ver])) {
          if (!agg[dataType][ver][variant]) agg[dataType][ver][variant] = {};
          for (const subj of Object.keys(dt[ver][variant])) {
            if (!agg[dataType][ver][variant][subj]) agg[dataType][ver][variant][subj] = {};
            for (const grade of Object.keys(dt[ver][variant][subj])) {
              agg[dataType][ver][variant][subj][grade] =
                (agg[dataType][ver][variant][subj][grade] || 0) + dt[ver][variant][subj][grade];
            }
          }
        }
      }
    }
  }
  return agg;
}

// =================================================================
// Build a statistics summary sheet with one or more aggregation blocks
// blocks: [{ label, aggData }, ...]
// =================================================================
async function buildStatsSheet(ws, title, blocks) {
  let row = 1;

  // Title
  ws.mergeCells(row, 1, row, 8);
  ws.getCell(row, 1).value = title;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = CENTER;
  ws.mergeCells(row, 9, row, TOTAL_COLS);
  ws.getCell(row, 9).value = '114.4.26~114.9.25';
  ws.getCell(row, 9).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 9).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 9).alignment = { horizontal: 'right', vertical: 'middle' };
  for (let c = 1; c <= TOTAL_COLS; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) cell.fill = solidFill(COLORS.titleBg);
    cell.border = THIN_BORDER;
  }
  row++;

  function writeAggBlock(label, aggData) {
    const examData = aggData['考卷'] || {};
    const whiteData = aggData['白卷'] || {};

    // Section title row
    ws.mergeCells(row, 1, row, TOTAL_COLS);
    ws.getCell(row, 1).value = label;
    ws.getCell(row, 1).font = { name: '微軟正黑體', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getCell(row, 1).fill = solidFill('FF2F5496');
    ws.getCell(row, 1).alignment = CENTER;
    for (let c = 1; c <= TOTAL_COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
    row++;

    // Header rows
    ws.getCell(row, 1).value = '品項';
    ws.getCell(row, 2).value = '卷別';
    ws.getCell(row, 3).value = '科目';
    let col = 4;
    for (const subj of SUBJECTS) {
      ws.mergeCells(row, col, row, col + 2);
      ws.getCell(row, col).value = subj;
      ws.getCell(row, col).alignment = CENTER;
      col += 3;
    }
    ws.getCell(row, col).value = '訂書';
    ws.getCell(row, col + 1).value = '退貨';
    ws.getCell(row, col + 2).value = '實際出貨';
    for (let c = 1; c <= TOTAL_COLS; c++) {
      styleCell(ws, row, c, { fill: solidFill(COLORS.headerBg), font: FONT_HEADER, border: THIN_BORDER, alignment: CENTER });
    }
    ws.mergeCells(row, 1, row + 1, 1);
    ws.mergeCells(row, 2, row + 1, 2);
    ws.mergeCells(row, 3, row + 1, 3);
    ws.mergeCells(row, col, row + 1, col);
    ws.mergeCells(row, col + 1, row + 1, col + 1);
    ws.mergeCells(row, col + 2, row + 1, col + 2);
    row++;

    col = 4;
    for (const subj of SUBJECTS) {
      for (const grade of GRADES) { ws.getCell(row, col).value = grade; col++; }
    }
    for (let c = 1; c <= TOTAL_COLS; c++) {
      styleCell(ws, row, c, { fill: solidFill(COLORS.subHeaderBg), font: FONT_SUBHEADER, border: THIN_BORDER, alignment: CENTER });
    }
    row++;

    function writeRow(vLabel, varLabel, areaLabel, rData, bgColor) {
      ws.getCell(row, 1).value = vLabel;
      ws.getCell(row, 2).value = varLabel;
      ws.getCell(row, 3).value = areaLabel;
      let rowTotal = 0, c = 4;
      for (const subj of SUBJECTS) {
        for (const grade of GRADES) {
          const qty = (rData[subj] && rData[subj][grade]) || 0;
          ws.getCell(row, c).value = qty || '';
          if (qty > 0) ws.getCell(row, c).numFmt = '#,##0';
          rowTotal += qty; c++;
        }
      }
      ws.getCell(row, c).value = rowTotal || 0; ws.getCell(row, c).numFmt = '#,##0';
      ws.getCell(row, c + 1).value = '';
      ws.getCell(row, c + 2).value = rowTotal || 0; ws.getCell(row, c + 2).numFmt = '#,##0';
      for (let cc = 1; cc <= TOTAL_COLS; cc++) {
        styleCell(ws, row, cc, { fill: solidFill(bgColor), font: FONT_DATA, border: THIN_BORDER, alignment: CENTER });
      }
      for (let sc = c; sc <= c + 2; sc++) {
        ws.getCell(row, sc).fill = solidFill(COLORS.summaryColBg);
        ws.getCell(row, sc).font = { ...FONT_DATA, bold: true };
      }
      const r = row; row++; return r;
    }

    function writeTotRow(lbl, area, datasets, bg, fnt, bdr) {
      ws.getCell(row, 1).value = '';
      ws.getCell(row, 2).value = lbl;
      ws.getCell(row, 3).value = area;
      let gt = 0, c = 4;
      for (const subj of SUBJECTS) {
        for (const grade of GRADES) {
          let t = 0;
          for (const ds of datasets) t += (ds[subj] && ds[subj][grade]) || 0;
          ws.getCell(row, c).value = t || 0; ws.getCell(row, c).numFmt = '#,##0';
          gt += t; c++;
        }
      }
      ws.getCell(row, c).value = gt || 0; ws.getCell(row, c).numFmt = '#,##0';
      ws.getCell(row, c + 1).value = '';
      ws.getCell(row, c + 2).value = gt || 0; ws.getCell(row, c + 2).numFmt = '#,##0';
      for (let cc = 1; cc <= TOTAL_COLS; cc++) {
        styleCell(ws, row, cc, { fill: solidFill(bg), font: fnt, border: bdr, alignment: CENTER });
      }
      for (let sc = c; sc <= c + 2; sc++) {
        ws.getCell(row, sc).fill = solidFill(COLORS.summaryColBg);
        ws.getCell(row, sc).font = { ...fnt, bold: true };
      }
      const r = row; row++; return r;
    }

    // K版
    const kD = examData['K'] || {}, kW = whiteData['K'] || {};
    const kA = kD['a'] || {}, kB = kD['b'] || {};
    const kWA = kW['a'] || {}, kWB = kW['b'] || {};
    const kStart = row;
    writeRow('K版', 'a卷', label, kA, COLORS.kVersionBg);
    writeRow('', 'b卷', label, kB, COLORS.kVersionBg);
    const kTot = writeTotRow('總數', '總數', [kA, kB, kWA, kWB], COLORS.totalBg, FONT_TOTAL, THIN_BORDER);
    ws.mergeCells(kStart, 1, kTot, 1);
    ws.getCell(kStart, 1).font = FONT_SECTION_LABEL;
    ws.getCell(kStart, 1).fill = solidFill(COLORS.kSectionColor);
    ws.getCell(kStart, 1).alignment = CENTER;

    // N版
    const nD = examData['N'] || {}, nW = whiteData['N'] || {};
    const nAll = nD['合計'] || {}, nWAll = nW['合計'] || {};
    const nMerged = {};
    for (const s of SUBJECTS) { nMerged[s] = {}; for (const g of GRADES) { nMerged[s][g] = ((nAll[s]&&nAll[s][g])||0)+((nWAll[s]&&nWAll[s][g])||0); } }
    const nR = writeRow('N版', '合計', label, nMerged, COLORS.nVersionBg);
    ws.getCell(nR, 1).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FF000000' } };
    ws.getCell(nR, 1).fill = solidFill(COLORS.nSectionColor);
    ws.getCell(nR, 1).alignment = CENTER;

    // H版
    const hD = examData['H'] || {}, hW = whiteData['H'] || {};
    const hA = hD['a'] || {}, hB = hD['b'] || {};
    const hWA = hW['a'] || {}, hWB = hW['b'] || {};
    const hStart = row;
    writeRow('H版', 'a卷', label, hA, COLORS.hVersionBg);
    writeRow('', 'b卷', label, hB, COLORS.hVersionBg);
    const hTot = writeTotRow('總數', '總數', [hA, hB, hWA, hWB], COLORS.totalBg, FONT_TOTAL, THIN_BORDER);
    ws.mergeCells(hStart, 1, hTot, 1);
    ws.getCell(hStart, 1).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FF000000' } };
    ws.getCell(hStart, 1).fill = solidFill(COLORS.hSectionColor);
    ws.getCell(hStart, 1).alignment = CENTER;

    // Grand total
    const allDS = [];
    for (const ver of ['K', 'N', 'H']) {
      for (const v of Object.values(examData[ver] || {})) allDS.push(v);
      for (const v of Object.values(whiteData[ver] || {})) allDS.push(v);
    }
    writeTotRow('', '總數', allDS, COLORS.totalBg, FONT_TOTAL, MEDIUM_BORDER_BOTTOM);
    row++; // spacing
  }

  // Write all blocks
  for (const block of blocks) {
    writeAggBlock(block.label, block.aggData);
  }

  // Apply borders
  for (let r = 1; r <= row; r++) {
    for (let c = 1; c <= TOTAL_COLS; c++) {
      const cell = ws.getCell(r, c);
      if (!cell.border || !cell.border.top) cell.border = THIN_BORDER;
    }
  }

  // Column widths
  ws.getColumn(1).width = 8;
  ws.getColumn(2).width = 8;
  ws.getColumn(3).width = 14;
  for (let c = 4; c <= TOTAL_COLS; c++) ws.getColumn(c).width = 7;
  ws.getColumn(TOTAL_COLS - 2).width = 8;
  ws.getColumn(TOTAL_COLS - 1).width = 6;
  ws.getColumn(TOTAL_COLS).width = 8;
  ws.getRow(1).height = 32;
}

// =================================================================
// Main: Generate all files
// =================================================================
async function main() {
  const salesData = buildExamData(allRecords);

  const outputDir = path.join(__dirname, 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Generate per-salesperson files (with customer summary sheet)
  for (const salesRep of targetSales) {
    const customerData = salesData[salesRep];
    if (!customerData || Object.keys(customerData).length === 0) {
      console.log(`${salesRep}: No data, skipping.`);
      continue;
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = '金安出版社';
    wb.created = new Date();

    // Add summary sheet FIRST (aggregated across all customers)
    const summaryWs = wb.addWorksheet('客戶加總', {
      properties: { defaultColWidth: 7 },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    const aggData = aggregateCustomerData(customerData);
    await buildStatsSheet(summaryWs, `114年上期 單冊考卷 — ${salesRep} 客戶加總`, [
      { label: `${salesRep} 合計`, aggData }
    ]);

    // Add individual customer sheets
    const customers = Object.keys(customerData).sort();
    for (const customer of customers) {
      let sheetName = customer.replace(/[\\\/\?\*\[\]]/g, '');
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

      const ws = wb.addWorksheet(sheetName, {
        properties: { defaultColWidth: 7 },
        pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      });

      await buildStyledSheet(ws, customer, customerData[customer]);
    }

    const outputPath = path.join(outputDir, `${salesRep}_114上單冊考卷.xlsx`);
    await wb.xlsx.writeFile(outputPath);
    console.log(`✓ ${outputPath} (${wb.worksheets.length} sheets, incl. 客戶加總)`);
  }

  // Generate summary file (all sales reps combined)
  const summaryWb = new ExcelJS.Workbook();
  summaryWb.creator = '金安出版社';
  summaryWb.created = new Date();

  // Build consolidated statistics sheet
  const statsWs = summaryWb.addWorksheet('全部統計', {
    properties: { defaultColWidth: 7 },
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  const statsBlocks = [];
  for (const salesRep of targetSales) {
    if (!salesData[salesRep] || Object.keys(salesData[salesRep]).length === 0) continue;
    statsBlocks.push({ label: salesRep, aggData: aggregateCustomerData(salesData[salesRep]) });
  }
  // Grand total
  const allCustomerDataMerged = {};
  for (const sr of targetSales) {
    if (!salesData[sr]) continue;
    for (const [cust, data] of Object.entries(salesData[sr])) {
      allCustomerDataMerged[`${sr}|${cust}`] = data;
    }
  }
  statsBlocks.push({ label: '北區合計', aggData: aggregateCustomerData(allCustomerDataMerged) });
  await buildStatsSheet(statsWs, '114年上期 單冊考卷 — 北區全部統計', statsBlocks);

  // Individual customer sheets
  for (const salesRep of targetSales) {
    const customerData = salesData[salesRep];
    if (!customerData || Object.keys(customerData).length === 0) continue;

    const customers = Object.keys(customerData).sort();
    for (const customer of customers) {
      let sheetName = `${salesRep}-${customer}`.replace(/[\\\/\?\*\[\]]/g, '');
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

      const ws = summaryWb.addWorksheet(sheetName, {
        properties: { defaultColWidth: 7 },
        pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      });

      await buildStyledSheet(ws, customer, customerData[customer]);
    }
  }

  const summaryPath = path.join(outputDir, '北區_114上單冊考卷_總表.xlsx');
  await summaryWb.xlsx.writeFile(summaryPath);
  console.log(`✓ ${summaryPath} (${summaryWb.worksheets.length} sheets)`);

  console.log('\nDone!');
}

main().catch(console.error);

