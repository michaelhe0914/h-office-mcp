// Generate beautifully styled xlsx files using exceljs
// Each sales rep gets one xlsx with customer sheets, with color blocks, borders, fonts

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results from both North and Central regions
const northRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_north_single_volume.json'), 'utf-8'));
let centralRecords = [];
try {
  centralRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_central_single_volume.json'), 'utf-8'));
} catch (e) {
  console.log("Central region query results not found or error reading, skipping central records.");
}
const allRecords = [...northRecords, ...centralRecords];

const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉瑋", "朱鵬學", "蔡榮訓", "曹原菘", "林廣強"];

// =================================================================
// Color Palette & Style Definitions
// =================================================================
const COLORS = {
  titleBg:     'FF1F4E79',  // 深藍
  titleFont:   'FFFFFFFF',  // 白
  headerBg:    'FF4472C4',  // 中藍 (section header row 1)
  headerFont:  'FFFFFFFF',
  subHeaderBg: 'FFD6E4F0',  // 淺藍 (grade row)
  subHeaderFont:'FF1F4E79',
  sxColor:     'FF70AD47',  // 雙向=綠
  sxLightBg:   'FFE2EFDA',
  q735Color:   'FFED7D31',  // 735=橘
  q735LightBg: 'FFFCE4D6',
  stpColor:    'FF5B9BD5',  // 試題篇=藍
  stpLightBg:  'FFDAEEF3',
  xlColor:     'FF7030A0',  // 新講義=紫
  xlLightBg:   'FFE8D5F5',
  totalBg:     'FFFFF2CC',  // 合計=淡黃
  totalFont:   'FF833C0B',
  dataBgEven:  'FFF2F2F2',  // 偶數行灰底
  dataBgOdd:   'FFFFFFFF',
  borderColor: 'FF000000',
  summaryColBg:'FFDCE6F1',  // 訂書/退貨/實際出貨欄
  nVersionBg:  'FFD9E2F3',  // N版淡藍
  kVersionBg:  'FFDCE6D0',  // K版淡綠
  hVersionBg:  'FFFBE5D6',  // H版淡橘
};

const FONT_TITLE = { name: '微軟正黑體', size: 14, bold: true, color: { argb: COLORS.titleFont } };
const FONT_HEADER = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.headerFont } };
const FONT_SUBHEADER = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.subHeaderFont } };
const FONT_DATA = { name: '微軟正黑體', size: 10 };
const FONT_TOTAL = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.totalFont } };
const FONT_SECTION_LABEL = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
const FONT_DATE = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FF1F4E79' } };

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
// Product parsing (same as before)
// =================================================================
function parseProduct(product, productClass) {
  const volMatch = product.match(/[（(]([１-６1-6])[）)]/);
  if (!volMatch) return null;
  const volMap = { '１': 1, '２': 2, '３': 3, '４': 4, '５': 5, '６': 6 };
  const volume = volMap[volMatch[1]] || parseInt(volMatch[1]);
  
  let version;
  if (productClass.includes('康版')) version = 'K';
  else if (productClass.includes('南版')) version = 'N';
  else if (productClass.includes('翰版')) version = 'H';
  else version = '綜合';
  
  let category;
  if (productClass.includes('雙向')) category = '雙向';
  else if (productClass.includes('735')) category = '735';
  else if (productClass.includes('試題')) category = '試題篇';
  else if (productClass.includes('新講義')) category = '新講義';
  else return null;
  
  let subject;
  if (category === '雙向') {
    if (product.includes('國文')) subject = '國文';
    else if (product.includes('英語')) subject = '英語';
    else if (product.includes('數學')) subject = '數學';
    else if (product.includes('自然')) subject = '自然';
    else return null;
  } else if (category === '735') {
    if (product.includes('地理')) subject = '地理';
    else if (product.includes('歷史')) subject = '歷史';
    else if (product.includes('國文')) subject = '國文';
    else if (product.includes('英語')) subject = '英語';
    else if (product.includes('數學')) subject = '數學';
    else if (product.includes('自然')) subject = '自然';
    else return null;
  } else if (category === '試題篇') {
    if (product.includes('英語文法')) subject = '英語';
    else if (product.includes('英語閱讀素養')) subject = '閱讀英文';
    else if (product.includes('數學')) subject = '數學';
    else return null;
  } else if (category === '新講義') {
    if (product.includes('數學')) subject = '數學';
    else if (product.includes('自然')) subject = '自然';
    else return null;
  }
  
  let grade;
  if (volume <= 2) grade = '一';
  else if (volume <= 4) grade = '三';
  else grade = '五';
  
  return { category, subject, volume, version, grade };
}

function buildSalesData(records) {
  const salesData = {};
  for (const r of records) {
    const parsed = parseProduct(r.product, r.productClass);
    if (!parsed) continue;
    const { category, subject, version, grade } = parsed;
    if (!salesData[r.sales]) salesData[r.sales] = {};
    if (!salesData[r.sales][r.customer]) salesData[r.sales][r.customer] = {};
    if (!salesData[r.sales][r.customer][category]) salesData[r.sales][r.customer][category] = {};
    if (!salesData[r.sales][r.customer][category][version]) salesData[r.sales][r.customer][category][version] = {};
    if (!salesData[r.sales][r.customer][category][version][subject]) salesData[r.sales][r.customer][category][version][subject] = {};
    salesData[r.sales][r.customer][category][version][subject][grade] = 
      (salesData[r.sales][r.customer][category][version][subject][grade] || 0) + r.qty;
  }
  return salesData;
}

// =================================================================
// Helper: apply style to a range of cells
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

function styleRange(ws, r1, c1, r2, c2, styles) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      styleCell(ws, r, c, styles);
    }
  }
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

function writeSummaryTable(ws, startRow, sxTotal, q735Total, stpTotal, xlTotal) {
  let row = startRow;
  
  // Title row
  ws.mergeCells(row, 1, row, 5);
  ws.getCell(row, 1).value = '各品項及年級加總表';
  ws.getCell(row, 1).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = CENTER;
  row++;
  
  // Header row
  const headers = ['品項', '一年級', '三年級', '五年級', '總本數'];
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    ws.getCell(row, c).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getCell(row, c).fill = solidFill(COLORS.headerBg);
    ws.getCell(row, c).alignment = CENTER;
  }
  row++;
  
  // Helper to generate formula for summing columns of a specific grade in a section total row
  // Grade is 0 (一), 1 (三), 2 (五)
  function getGradeSumFormula(totalRow, numSubjects, gradeIndex) {
    const cols = [];
    for (let s = 0; s < numSubjects; s++) {
      const colNum = 4 + s * 3 + gradeIndex;
      cols.push(`${colLetter(colNum)}${totalRow}`);
    }
    return `SUM(${cols.join(',')})`;
  }
  
  // 雙向 (4 subjects, sxTotal)
  ws.getCell(row, 1).value = '雙向';
  ws.getCell(row, 2).value = { formula: getGradeSumFormula(sxTotal, 4, 0) };
  ws.getCell(row, 3).value = { formula: getGradeSumFormula(sxTotal, 4, 1) };
  ws.getCell(row, 4).value = { formula: getGradeSumFormula(sxTotal, 4, 2) };
  ws.getCell(row, 5).value = { formula: `R${sxTotal}` }; // 實際出貨 column for 雙向 (4 subjects: col 18 is R)
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.sxLightBg);
  }
  row++;
  
  // 735 (6 subjects, q735Total)
  ws.getCell(row, 1).value = '735';
  ws.getCell(row, 2).value = { formula: getGradeSumFormula(q735Total, 6, 0) };
  ws.getCell(row, 3).value = { formula: getGradeSumFormula(q735Total, 6, 1) };
  ws.getCell(row, 4).value = { formula: getGradeSumFormula(q735Total, 6, 2) };
  ws.getCell(row, 5).value = { formula: `X${q735Total}` }; // 實際出貨 column for 735 (6 subjects: col 24 is X)
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.q735LightBg);
  }
  row++;
  
  // 試題篇 (3 subjects, stpTotal)
  ws.getCell(row, 1).value = '試題篇';
  ws.getCell(row, 2).value = { formula: getGradeSumFormula(stpTotal, 3, 0) };
  ws.getCell(row, 3).value = { formula: getGradeSumFormula(stpTotal, 3, 1) };
  ws.getCell(row, 4).value = { formula: getGradeSumFormula(stpTotal, 3, 2) };
  ws.getCell(row, 5).value = { formula: `O${stpTotal}` }; // 實際出貨 column for 試題篇 (3 subjects: col 15 is O)
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.stpLightBg);
  }
  row++;
  
  // 新講義 (2 subjects, xlTotal)
  ws.getCell(row, 1).value = '新講義';
  ws.getCell(row, 2).value = { formula: getGradeSumFormula(xlTotal, 2, 0) };
  ws.getCell(row, 3).value = { formula: getGradeSumFormula(xlTotal, 2, 1) };
  ws.getCell(row, 4).value = { formula: getGradeSumFormula(xlTotal, 2, 2) };
  ws.getCell(row, 5).value = { formula: `L${xlTotal}` }; // 實際出貨 column for 新講義 (2 subjects: col 12 is L)
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.xlLightBg);
  }
  row++;
  
  // 合計 Row
  const sumStartRow = row - 4; // 雙向 row
  const sumEndRow = row - 1;   // 新講義 row
  
  ws.getCell(row, 1).value = '合計';
  ws.getCell(row, 2).value = { formula: `SUM(B${sumStartRow}:B${sumEndRow})` };
  ws.getCell(row, 3).value = { formula: `SUM(C${sumStartRow}:C${sumEndRow})` };
  ws.getCell(row, 4).value = { formula: `SUM(D${sumStartRow}:D${sumEndRow})` };
  ws.getCell(row, 5).value = { formula: `SUM(E${sumStartRow}:E${sumEndRow})` };
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.totalBg);
    ws.getCell(row, c).font = FONT_TOTAL;
  }
  row++;
  
  // Style and format all cells in the summary table
  const tableStart = startRow;
  const tableEnd = row - 1;
  for (let r = tableStart; r <= tableEnd; r++) {
    for (let c = 1; c <= 5; c++) {
      const cell = ws.getCell(r, c);
      cell.border = THIN_BORDER;
      cell.alignment = CENTER;
      
      if (r !== tableStart && r !== tableEnd && c > 1) {
        cell.font = FONT_DATA;
      } else if (r !== tableStart && r !== tableEnd && c === 1) {
        cell.font = { ...FONT_DATA, bold: true };
      }
      
      if (r >= tableStart + 2 && c >= 2) {
        cell.numFmt = '#,##0';
      }
    }
  }
  
  return row;
}

// =================================================================
// Build a styled sheet for one customer
// =================================================================
async function buildStyledSheet(ws, customer, data) {
  const grades = ['一', '三', '五'];
  let row = 1;

  // ==========================================
  // Title Row
  // ==========================================
  ws.mergeCells(row, 1, row, 6);
  ws.getCell(row, 1).value = '114年上期單冊講義';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = { ...CENTER };

  ws.mergeCells(row, 7, row, 10);
  ws.getCell(row, 7).value = customer;
  ws.getCell(row, 7).font = { ...FONT_TITLE, size: 12 };
  ws.getCell(row, 7).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 7).alignment = CENTER;

  ws.mergeCells(row, 11, row, 16);
  ws.getCell(row, 11).value = '114.4.26~114.9.25';
  ws.getCell(row, 11).font = { name: '微軟正黑體', size: 11, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 11).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 11).alignment = { horizontal: 'right', vertical: 'middle' };

  // Fill remaining title row cells with dark bg + border
  for (let c = 1; c <= 24; c++) {
    const cell = ws.getCell(row, c);
    if (!cell.fill || !cell.fill.fgColor) {
      cell.fill = solidFill(COLORS.titleBg);
    }
    cell.border = THIN_BORDER;
  }
  row += 1;

  // Helper to write a section
  function writeSection(sectionName, sectionLabel, subjects, versions, sectionColor, sectionLightBg, sectionData) {
    let startRow = row;
    const numSubjCols = subjects.length * 3; // each subject has 3 grade columns
    const totalCols = 3 + numSubjCols + 3; // product+version+period + subjects + 訂書+退貨+實際

    // ---- Header Row 1: Section header with subject groups ----
    ws.getCell(row, 1).value = '產品';
    ws.getCell(row, 2).value = '版本';
    ws.getCell(row, 3).value = customer;
    
    let col = 4;
    for (const subj of subjects) {
      ws.mergeCells(row, col, row, col + 2);
      ws.getCell(row, col).value = subj;
      ws.getCell(row, col).alignment = CENTER;
      col += 3;
    }
    
    // Summary columns
    ws.getCell(row, col).value = '訂書';
    ws.getCell(row, col + 1).value = '退貨';
    ws.getCell(row, col + 2).value = '實際出貨';

    // Style header row 1
    for (let c = 1; c <= totalCols; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(sectionColor),
        font: FONT_HEADER,
        border: THIN_BORDER,
        alignment: CENTER,
      });
    }
    // Merge header cells
    ws.mergeCells(row, 1, row + 1, 1);
    ws.mergeCells(row, 2, row + 1, 2);
    ws.mergeCells(row, 3, row + 1, 3);
    row++;

    // ---- Header Row 2: Grade labels ----
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = '';
    ws.getCell(row, 3).value = '';
    col = 4;
    for (const subj of subjects) {
      for (const grade of grades) {
        ws.getCell(row, col).value = grade;
        col++;
      }
    }
    ws.getCell(row, col).value = '';
    ws.getCell(row, col + 1).value = '';
    ws.getCell(row, col + 2).value = '';

    for (let c = 1; c <= totalCols; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(COLORS.subHeaderBg),
        font: FONT_SUBHEADER,
        border: THIN_BORDER,
        alignment: CENTER,
      });
    }
    row++;

    // ---- Data Rows (one per version) ----
    const versionBgMap = { 'N': COLORS.nVersionBg, 'K': COLORS.kVersionBg, 'H': COLORS.hVersionBg, '綜合': COLORS.xlLightBg };
    const dataStartRow = row;

    for (let vi = 0; vi < versions.length; vi++) {
      const v = versions[vi];
      const vLabel = v === 'K' ? 'K版' : v === 'N' ? 'N版' : v === 'H' ? 'H版' : '綜合版 合計';
      const vData = sectionData[v] || {};
      const bgColor = sectionLightBg;

      ws.getCell(row, 1).value = vi === 0 ? sectionLabel : '';
      ws.getCell(row, 2).value = vLabel;
      ws.getCell(row, 3).value = '114上';

      let rowTotal = 0;
      col = 4;
      for (const subj of subjects) {
        for (const grade of grades) {
          const qty = (vData[subj] && vData[subj][grade]) || 0;
          ws.getCell(row, col).value = qty || '';
          if (qty > 0) {
            ws.getCell(row, col).numFmt = '#,##0';
          }
          col++;
        }
        rowTotal += grades.reduce((s, g) => s + ((vData[subj] && vData[subj][g]) || 0), 0);
      }

      ws.getCell(row, col).value = rowTotal || 0;
      ws.getCell(row, col).numFmt = '#,##0';
      ws.getCell(row, col + 1).value = '';
      ws.getCell(row, col + 2).value = rowTotal || 0;
      ws.getCell(row, col + 2).numFmt = '#,##0';

      // Styling for data row
      const rowBg = versionBgMap[v] || bgColor;
      for (let c = 1; c <= totalCols; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(rowBg),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c <= 3 ? LEFT : CENTER,
        });
      }
      
      // Section label styling
      if (vi === 0) {
        ws.getCell(row, 1).font = FONT_SECTION_LABEL;
        ws.getCell(row, 1).fill = solidFill(sectionColor);
        ws.getCell(row, 1).alignment = CENTER;
      }
      
      // Version label bold
      ws.getCell(row, 2).font = { ...FONT_DATA, bold: true };
      ws.getCell(row, 2).alignment = CENTER;
      ws.getCell(row, 3).alignment = CENTER;

      // Summary columns highlight
      for (let sc = col; sc <= col + 2; sc++) {
        ws.getCell(row, sc).fill = solidFill(COLORS.summaryColBg);
        ws.getCell(row, sc).font = { ...FONT_DATA, bold: true };
      }

      row++;
    }

    // Merge section label vertically
    if (versions.length > 1) {
      ws.mergeCells(dataStartRow, 1, dataStartRow + versions.length - 1, 1);
    }

    // ---- Blank separator row (with borders) ----
    for (let c = 1; c <= totalCols; c++) {
      styleCell(ws, row, c, { border: THIN_BORDER });
    }
    row++;

    // ---- 合計 Row ----
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = '合計';
    ws.getCell(row, 3).value = '114上';

    let grandTotal = 0;
    col = 4;
    for (const subj of subjects) {
      for (const grade of grades) {
        let total = 0;
        for (const v of versions) {
          total += ((sectionData[v] || {})[subj] || {})[grade] || 0;
        }
        ws.getCell(row, col).value = total || 0;
        ws.getCell(row, col).numFmt = '#,##0';
        grandTotal += total;
        col++;
      }
    }

    ws.getCell(row, col).value = grandTotal || 0;
    ws.getCell(row, col).numFmt = '#,##0';
    ws.getCell(row, col + 1).value = 0;
    ws.getCell(row, col + 2).value = grandTotal || 0;
    ws.getCell(row, col + 2).numFmt = '#,##0';

    // Style 合計 row
    for (let c = 1; c <= totalCols; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(COLORS.totalBg),
        font: FONT_TOTAL,
        border: MEDIUM_BORDER_BOTTOM,
        alignment: c <= 3 ? CENTER : CENTER,
      });
    }

    const totalRow = row;
    row += 2; // spacing after section
    return { nextRow: row, totalRow };
  }

  // ==========================================
  // Section 1: 雙向
  // ==========================================
  let res = writeSection('雙向', '雙向', ['國文', '英語', '數學', '自然'],
    ['N', 'K', 'H'], COLORS.sxColor, COLORS.sxLightBg, data['雙向'] || {});
  row = res.nextRow;
  const sxTotalRow = res.totalRow;

  // ==========================================
  // Section 2: 735
  // ==========================================
  res = writeSection('735', '735', ['國文', '英語', '數學', '自然', '地理', '歷史'],
    ['N', 'K', 'H'], COLORS.q735Color, COLORS.q735LightBg, data['735'] || {});
  row = res.nextRow;
  const q735TotalRow = res.totalRow;

  // ==========================================
  // Section 3: 試題篇
  // ==========================================
  res = writeSection('試題篇', '試題篇', ['英語', '閱讀英文', '數學'],
    ['K', 'N', 'H'], COLORS.stpColor, COLORS.stpLightBg, data['試題篇'] || {});
  row = res.nextRow;
  const stpTotalRow = res.totalRow;

  // ==========================================
  // Section 4: 新講義
  // ==========================================
  res = writeSection('新講義', '新講義', ['數學', '自然'],
    ['綜合'], COLORS.xlColor, COLORS.xlLightBg, data['新講義'] || {});
  row = res.nextRow;
  const xlTotalRow = res.totalRow;

  // ==========================================
  // Summary Table: 各品項及年級加總表
  // ==========================================
  const startSumRow = row;
  row = writeSummaryTable(ws, startSumRow, sxTotalRow, q735TotalRow, stpTotalRow, xlTotalRow);

  // ==========================================
  // Final pass: apply borders to ALL cells in the used range
  // This catches merged-cell slaves, blank rows, and any gaps
  // ==========================================
  const lastRow = row;
  const lastCol = 24; // max possible columns across all sections
  for (let r = 1; r <= lastRow; r++) {
    const limitCol = (r >= startSumRow) ? 5 : lastCol;
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
  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 8;
  for (let c = 4; c <= 24; c++) {
    ws.getColumn(c).width = 7;
  }

  // Row heights
  ws.getRow(1).height = 28;
}

// =================================================================
// Helper: sum all customers' data for a sales rep
// =================================================================
function sumRepresentativeData(customerData) {
  const summary = {};
  for (const customer of Object.keys(customerData)) {
    const data = customerData[customer];
    for (const category of Object.keys(data)) {
      if (!summary[category]) summary[category] = {};
      for (const version of Object.keys(data[category])) {
        if (!summary[category][version]) summary[category][version] = {};
        for (const subject of Object.keys(data[category][version])) {
          if (!summary[category][version][subject]) summary[category][version][subject] = {};
          for (const grade of Object.keys(data[category][version][subject])) {
            summary[category][version][subject][grade] = 
              (summary[category][version][subject][grade] || 0) + data[category][version][subject][grade];
          }
        }
      }
    }
  }
  return summary;
}

// =================================================================
// Helper: sum sales reps' data for a specific list of reps
// =================================================================
function sumRegionTotalData(salesData, repsList) {
  const summary = {};
  for (const salesRep of repsList) {
    const repData = sumRepresentativeData(salesData[salesRep] || {});
    for (const category of Object.keys(repData)) {
      if (!summary[category]) summary[category] = {};
      for (const version of Object.keys(repData[category])) {
        if (!summary[category][version]) summary[category][version] = {};
        for (const subject of Object.keys(repData[category][version])) {
          if (!summary[category][version][subject]) summary[category][version][subject] = {};
          for (const grade of Object.keys(repData[category][version][subject])) {
            summary[category][version][subject][grade] = 
              (summary[category][version][subject][grade] || 0) + repData[category][version][subject][grade];
          }
        }
      }
    }
  }
  return summary;
}

const northSales = ["何光傑", "李敏豪", "林智偉", "康晉瑋", "朱鵬學"];
const centralSales = ["蔡榮訓", "曹原菘", "林廣強"];

// =================================================================
// Main: Generate all files
// =================================================================
async function main() {
  const salesData = buildSalesData(allRecords);

  const outputDir = path.join(__dirname, 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // 1. Generate files for each sales rep (with "客戶加總" as the first sheet)
  for (const salesRep of targetSales) {
    const customerData = salesData[salesRep];
    if (!customerData || Object.keys(customerData).length === 0) {
      console.log(`${salesRep}: No data, skipping.`);
      continue;
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = '金安出版社';
    wb.created = new Date();

    // Calculate customer summary total data
    const repSummaryData = sumRepresentativeData(customerData);

    // Create the "客戶加總" sheet first
    const wsSummary = wb.addWorksheet('客戶加總', {
      properties: { defaultColWidth: 7 },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    await buildStyledSheet(wsSummary, '客戶加總', repSummaryData);

    // Sort customer names
    const customers = Object.keys(customerData).sort();

    // Create individual customer sheets
    for (const customer of customers) {
      let sheetName = customer.replace(/[\\\/?*\[\]]/g, '');
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

      const ws = wb.addWorksheet(sheetName, {
        properties: { defaultColWidth: 7 },
        pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      });

      await buildStyledSheet(ws, customer, customerData[customer]);
    }

    const outputPath = path.join(outputDir, `${salesRep}_114上單冊講義.xlsx`);
    await wb.xlsx.writeFile(outputPath);
    console.log(`✓ ${outputPath} (${customers.length + 1} sheets)`);
  }

  // 2. Delete outdated national grand summary if exists
  const oldGrandSummaryPath = path.join(outputDir, '單冊講義總數量合計總表.xlsx');
  if (fs.existsSync(oldGrandSummaryPath)) {
    fs.unlinkSync(oldGrandSummaryPath);
    console.log(`Deleted outdated combined file: ${oldGrandSummaryPath}`);
  }

  // 3. Generate the North region grand summary file "北區單冊講義總數量合計總表.xlsx"
  console.log("\nGenerating North region grand summary table...");
  const northTotalData = sumRegionTotalData(salesData, northSales);

  const wbNorth = new ExcelJS.Workbook();
  wbNorth.creator = '金安出版社';
  wbNorth.created = new Date();

  const wsNorth = wbNorth.addWorksheet('總數量合計總表', {
    properties: { defaultColWidth: 7 },
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  await buildStyledSheet(wsNorth, '北區總數量合計總表', northTotalData);

  const northSummaryPath = path.join(outputDir, '北區單冊講義總數量合計總表.xlsx');
  await wbNorth.xlsx.writeFile(northSummaryPath);
  console.log(`✓ ${northSummaryPath} (1 sheet)`);

  // 4. Generate the Central region grand summary file "中區單冊講義總數量合計總表.xlsx"
  console.log("\nGenerating Central region grand summary table...");
  const centralTotalData = sumRegionTotalData(salesData, centralSales);

  const wbCentral = new ExcelJS.Workbook();
  wbCentral.creator = '金安出版社';
  wbCentral.created = new Date();

  const wsCentralSheet = wbCentral.addWorksheet('總數量合計總表', {
    properties: { defaultColWidth: 7 },
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  await buildStyledSheet(wsCentralSheet, '中區總數量合計總表', centralTotalData);

  const centralSummaryPath = path.join(outputDir, '中區單冊講義總數量合計總表.xlsx');
  await wbCentral.xlsx.writeFile(centralSummaryPath);
  console.log(`✓ ${centralSummaryPath} (1 sheet)`);

  console.log('\nDone!');
}

main().catch(console.error);
