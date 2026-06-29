// Generate beautifully styled xlsx files using exceljs
// Each sales rep gets one xlsx with customer sheets, with color blocks, borders, fonts
// For 中區 sales reps: 蔡榮訓、曹原菘、林廣強

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_central_single_volume.json'), 'utf-8'));
const targetSales = ["蔡榮訓", "曹原菘", "林廣強"];

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
// Product parsing
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
// Helper: apply style to cells
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

    row += 2; // spacing after section
    return row;
  }

  // ==========================================
  // Section 1: 雙向
  // ==========================================
  row = writeSection('雙向', '雙向', ['國文', '英語', '數學', '自然'],
    ['N', 'K', 'H'], COLORS.sxColor, COLORS.sxLightBg, data['雙向'] || {});

  // ==========================================
  // Section 2: 735
  // ==========================================
  row = writeSection('735', '735', ['國文', '英語', '數學', '自然', '地理', '歷史'],
    ['N', 'K', 'H'], COLORS.q735Color, COLORS.q735LightBg, data['735'] || {});

  // ==========================================
  // Section 3: 試題篇
  // ==========================================
  row = writeSection('試題篇', '試題篇', ['英語', '閱讀英文', '數學'],
    ['K', 'N', 'H'], COLORS.stpColor, COLORS.stpLightBg, data['試題篇'] || {});

  // ==========================================
  // Section 4: 新講義
  // ==========================================
  row = writeSection('新講義', '新講義', ['數學', '自然'],
    ['綜合'], COLORS.xlColor, COLORS.xlLightBg, data['新講義'] || {});

  // ==========================================
  // Final pass: apply borders to ALL cells in the used range
  // ==========================================
  const lastRow = row;
  const lastCol = 24; // max possible columns across all sections
  for (let r = 1; r <= lastRow; r++) {
    for (let c = 1; c <= lastCol; c++) {
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
// Main: Generate all files
// =================================================================
async function main() {
  const salesData = buildSalesData(allRecords);

  const outputDir = path.join(__dirname, 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  for (const salesRep of targetSales) {
    const customerData = salesData[salesRep];
    if (!customerData || Object.keys(customerData).length === 0) {
      console.log(`${salesRep}: No data, skipping.`);
      continue;
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = '金安出版社';
    wb.created = new Date();

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

    const outputPath = path.join(outputDir, `${salesRep}_114上單冊講義.xlsx`);
    await wb.xlsx.writeFile(outputPath);
    console.log(`✓ ${outputPath} (${customers.length} sheets)`);
  }

  console.log('\nDone!');
}

main().catch(console.error);
