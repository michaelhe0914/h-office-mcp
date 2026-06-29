// Generate beautifully styled xlsx files for regional textbook statistics using exceljs
// The sheet contains 4 sheets: 總表, 北區, 中區, 南區
// Each sheet has the standard 4 sections (雙向, 735, 試題篇, 新講義) and grade summary table

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results (includes all records for North, Central, South)
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_all_regions.json'), 'utf-8'));

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

function buildRegionSalesData(records) {
  const regionData = {
    '總表': {},
    '北區': {},
    '中區': {},
    '南區': {}
  };
  
  for (const r of records) {
    const parsed = parseProduct(r.product, r.productClass);
    if (!parsed) continue;
    
    const { category, subject, version, grade } = parsed;
    const zone = r.zone;
    const qty = r.qty;
    
    if (zone !== '北區' && zone !== '中區' && zone !== '南區') continue;
    
    // Add to specific region
    if (!regionData[zone][category]) regionData[zone][category] = {};
    if (!regionData[zone][category][version]) regionData[zone][category][version] = {};
    if (!regionData[zone][category][version][subject]) regionData[zone][category][version][subject] = {};
    regionData[zone][category][version][subject][grade] = 
      (regionData[zone][category][version][subject][grade] || 0) + qty;
      
    // Add to grand total sheet
    if (!regionData['總表'][category]) regionData['總表'][category] = {};
    if (!regionData['總表'][category][version]) regionData['總表'][category][version] = {};
    if (!regionData['總表'][category][version][subject]) regionData['總表'][category][version][subject] = {};
    regionData['總表'][category][version][subject][grade] = 
      (regionData['總表'][category][version][subject][grade] || 0) + qty;
  }
  
  return regionData;
}

// =================================================================
// Helpers: styling
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
// Build a styled sheet for one region
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
    const numSubjCols = subjects.length * 3;
    const totalCols = 3 + numSubjCols + 3;

    // Header Row 1
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
    
    ws.getCell(row, col).value = '訂書';
    ws.getCell(row, col + 1).value = '退貨';
    ws.getCell(row, col + 2).value = '實際出貨';

    for (let c = 1; c <= totalCols; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(sectionColor),
        font: FONT_HEADER,
        border: THIN_BORDER,
        alignment: CENTER,
      });
    }
    ws.mergeCells(row, 1, row + 1, 1);
    ws.mergeCells(row, 2, row + 1, 2);
    ws.mergeCells(row, 3, row + 1, 3);
    row++;

    // Header Row 2
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

    // Data Rows
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

      const rowBg = versionBgMap[v] || bgColor;
      for (let c = 1; c <= totalCols; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(rowBg),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c <= 3 ? LEFT : CENTER,
        });
      }
      
      if (vi === 0) {
        ws.getCell(row, 1).font = FONT_SECTION_LABEL;
        ws.getCell(row, 1).fill = solidFill(sectionColor);
        ws.getCell(row, 1).alignment = CENTER;
      }
      
      ws.getCell(row, 2).font = { ...FONT_DATA, bold: true };
      ws.getCell(row, 2).alignment = CENTER;
      ws.getCell(row, 3).alignment = CENTER;

      for (let sc = col; sc <= col + 2; sc++) {
        ws.getCell(row, sc).fill = solidFill(COLORS.summaryColBg);
        ws.getCell(row, sc).font = { ...FONT_DATA, bold: true };
      }

      row++;
    }

    if (versions.length > 1) {
      ws.mergeCells(dataStartRow, 1, dataStartRow + versions.length - 1, 1);
    }

    // Blank row
    for (let c = 1; c <= totalCols; c++) {
      styleCell(ws, row, c, { border: THIN_BORDER });
    }
    row++;

    // 合計 Row
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

    for (let c = 1; c <= totalCols; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(COLORS.totalBg),
        font: FONT_TOTAL,
        border: MEDIUM_BORDER_BOTTOM,
        alignment: CENTER,
      });
    }

    const totalRow = row;
    row += 2;
    return { nextRow: row, totalRow };
  }

  // Section 1: 雙向
  let res = writeSection('雙向', '雙向', ['國文', '英語', '數學', '自然'],
    ['N', 'K', 'H'], COLORS.sxColor, COLORS.sxLightBg, data['雙向'] || {});
  row = res.nextRow;
  const sxTotalRow = res.totalRow;

  // Section 2: 735
  res = writeSection('735', '735', ['國文', '英語', '數學', '自然', '地理', '歷史'],
    ['N', 'K', 'H'], COLORS.q735Color, COLORS.q735LightBg, data['735'] || {});
  row = res.nextRow;
  const q735TotalRow = res.totalRow;

  // Section 3: 試題篇
  res = writeSection('試題篇', '試題篇', ['英語', '閱讀英文', '數學'],
    ['K', 'N', 'H'], COLORS.stpColor, COLORS.stpLightBg, data['試題篇'] || {});
  row = res.nextRow;
  const stpTotalRow = res.totalRow;

  // Section 4: 新講義
  res = writeSection('新講義', '新講義', ['數學', '自然'],
    ['綜合'], COLORS.xlColor, COLORS.xlLightBg, data['新講義'] || {});
  row = res.nextRow;
  const xlTotalRow = res.totalRow;

  // Summary Table: 各品項及年級加總表
  const startSumRow = row;
  row = writeSummaryTable(ws, startSumRow, sxTotalRow, q735TotalRow, stpTotalRow, xlTotalRow);

  // Final pass: apply borders to ALL cells in the used range
  const lastRow = row;
  const lastCol = 24;
  for (let r = 1; r <= lastRow; r++) {
    const limitCol = (r >= startSumRow) ? 5 : lastCol;
    for (let c = 1; c <= limitCol; c++) {
      const cell = ws.getCell(r, c);
      if (!cell.border || !cell.border.top) {
        cell.border = THIN_BORDER;
      }
    }
  }

  // Column widths
  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 12; // wider for zone names like '總數量合計總表'
  for (let c = 4; c <= 24; c++) {
    ws.getColumn(c).width = 7;
  }

  // Row heights
  ws.getRow(1).height = 28;
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
  ws.getCell(row, 5).value = { formula: `R${sxTotal}` };
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.sxLightBg);
  }
  row++;
  
  // 735 (6 subjects, q735Total)
  ws.getCell(row, 1).value = '735';
  ws.getCell(row, 2).value = { formula: getGradeSumFormula(q735Total, 6, 0) };
  ws.getCell(row, 3).value = { formula: getGradeSumFormula(q735Total, 6, 1) };
  ws.getCell(row, 4).value = { formula: getGradeSumFormula(q735Total, 6, 2) };
  ws.getCell(row, 5).value = { formula: `X${q735Total}` };
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.q735LightBg);
  }
  row++;
  
  // 試題篇 (3 subjects, stpTotal)
  ws.getCell(row, 1).value = '試題篇';
  ws.getCell(row, 2).value = { formula: getGradeSumFormula(stpTotal, 3, 0) };
  ws.getCell(row, 3).value = { formula: getGradeSumFormula(stpTotal, 3, 1) };
  ws.getCell(row, 4).value = { formula: getGradeSumFormula(stpTotal, 3, 2) };
  ws.getCell(row, 5).value = { formula: `O${stpTotal}` };
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.stpLightBg);
  }
  row++;
  
  // 新講義 (2 subjects, xlTotal)
  ws.getCell(row, 1).value = '新講義';
  ws.getCell(row, 2).value = { formula: getGradeSumFormula(xlTotal, 2, 0) };
  ws.getCell(row, 3).value = { formula: getGradeSumFormula(xlTotal, 2, 1) };
  ws.getCell(row, 4).value = { formula: getGradeSumFormula(xlTotal, 2, 2) };
  ws.getCell(row, 5).value = { formula: `L${xlTotal}` };
  
  for (let c = 1; c <= 5; c++) {
    ws.getCell(row, c).fill = solidFill(COLORS.xlLightBg);
  }
  row++;
  
  // 合計 Row
  const sumStartRow = row - 4;
  const sumEndRow = row - 1;
  
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
// Main: Generate the combined workbook
// =================================================================
async function main() {
  const regionSalesData = buildRegionSalesData(allRecords);

  const outputDir = path.join(__dirname, 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = '金安出版社';
  wb.created = new Date();

  const sheetsOrder = ['總表', '北區', '中區', '南區'];

  for (const zone of sheetsOrder) {
    const ws = wb.addWorksheet(zone, {
      properties: { defaultColWidth: 7 },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    console.log(`Writing sheet for: ${zone}...`);
    await buildStyledSheet(ws, zone, regionSalesData[zone]);
  }

  const outputPath = path.join(outputDir, '各區單冊講義產品總數量表.xlsx');
  await wb.xlsx.writeFile(outputPath);
  console.log(`\n✓ Successfully generated: ${outputPath} (${sheetsOrder.length} sheets)`);
}

main().catch(console.error);
