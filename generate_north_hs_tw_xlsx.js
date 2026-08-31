import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// =================================================================
// Colors & Styles Definition
// =================================================================
const COLORS = {
  titleBg:         'FF1F4E79',  // 深藍
  headerBg:        'FF4472C4',  // 中藍
  subHeaderBg:     'FFD6E4F0',  // 淺藍
  textbookColor:   'FF70AD47',  // 課本 = 綠色
  textbookLightBg: 'FFE2EFDA',
  prepColor:       'FF5B9BD5',  // 備課 = 藍色
  prepLightBg:     'FFDAEEF3',
  totalBg:         'FFFFF2CC',  // 合計 = 淡黃色
  totalFont:       'FF833C0B',
  dataBgEven:      'FFF2F2F2',  // 偶數行灰底
  dataBgOdd:       'FFFFFFFF',
  borderColor:     'FF000000',
  negativeBg:      'FFFFC7CE',  // 負數提醒淡紅
};

const FONT_TITLE = { name: '微軟正黑體', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
const FONT_HEADER = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const FONT_DATA = { name: '微軟正黑體', size: 10 };
const FONT_BOLD = { name: '微軟正黑體', size: 10, bold: true };
const FONT_TOTAL = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.totalFont } };
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

// =================================================================
// Product Canonicals
// =================================================================
const TEXTBOOK_PRODUCTS = [
  '*高中職 閩南語文(１)',
  '*高中職 閩南語文(２)',
  '高中職 閩南語文(全一冊)上',
  '高中職 閩南語文(全一冊)下',
  '高中職 閩南語文(全一冊)乙版',
  '*高中職 閩南語文(全一冊)'
];

const PREP_PRODUCTS = [
  '*高中職 閩南語文(１)備課用書包'
];

const ALL_CANONICAL_PRODUCTS = [...TEXTBOOK_PRODUCTS, ...PREP_PRODUCTS];

function getDisplayName(prod) {
  return prod.replace(/^\*/, '');
}

// =================================================================
// Sheet Builders
// =================================================================
function buildCustomerSheet(ws, customer, zone, sales, custData) {
  let row = 1;
  const COLS = 7;
  
  // Title Row
  ws.mergeCells(row, 1, row, 4);
  ws.getCell(row, 1).value = '金安出版社 高中台語出退貨明細表';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  
  ws.mergeCells(row, 5, row, COLS);
  ws.getCell(row, 5).value = '2025/04/26 ~ 2025/10/25';
  ws.getCell(row, 5).font = FONT_DATE;
  ws.getCell(row, 5).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
  }
  ws.getRow(row).height = 30;
  row++;
  
  // Metadata Row
  ws.mergeCells(row, 1, row, 3);
  ws.getCell(row, 1).value = `客戶名稱：${customer}`;
  ws.getCell(row, 1).font = FONT_BOLD;
  ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  
  ws.mergeCells(row, 4, row, 5);
  ws.getCell(row, 4).value = `負責業務：${sales}`;
  ws.getCell(row, 4).font = FONT_BOLD;
  ws.getCell(row, 4).alignment = { horizontal: 'left', vertical: 'middle' };
  
  ws.mergeCells(row, 6, row, COLS);
  ws.getCell(row, 6).value = `所屬區域：${zone}`;
  ws.getCell(row, 6).font = FONT_BOLD;
  ws.getCell(row, 6).alignment = { horizontal: 'left', vertical: 'middle' };
  
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
    ws.getCell(row, c).fill = solidFill(COLORS.subHeaderBg);
  }
  ws.getRow(row).height = 22;
  row++;
  
  // Spacing
  ws.getRow(row).height = 10;
  row++;
  
  // Header Row
  const headers = ['序號', '產品類別', '產品名稱', '出貨數量', '退貨量', '實際出貨', '退書率'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: { horizontal: 'center', vertical: 'middle' }
    });
  }
  ws.getRow(row).height = 24;
  row++;
  
  const textbookStartRow = row;
  let idx = 1;
  
  // Textbook Product Rows
  for (let i = 0; i < TEXTBOOK_PRODUCTS.length; i++) {
    const prod = TEXTBOOK_PRODUCTS[i];
    const data = custData[prod] || { qty: 0, rtn_qty: 0 };
    const bgColor = i % 2 === 0 ? COLORS.dataBgOdd : COLORS.dataBgEven;
    
    ws.getCell(row, 1).value = idx++;
    ws.getCell(row, 2).value = '課本';
    ws.getCell(row, 3).value = getDisplayName(prod);
    ws.getCell(row, 4).value = data.qty;
    ws.getCell(row, 5).value = data.rtn_qty;
    ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
    ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
    
    ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 3).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
    
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).numFmt = '#,##0';
    ws.getCell(row, 7).numFmt = '0.0%';
    
    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER
      });
    }
    row++;
  }
  const textbookEndRow = row - 1;
  const textbookSubtotalRow = row;
  
  // Textbook Subtotal Row
  ws.getCell(row, 1).value = '課本 小計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = { formula: `=SUM(D${textbookStartRow}:D${textbookEndRow})` };
  ws.getCell(row, 5).value = { formula: `=SUM(E${textbookStartRow}:E${textbookEndRow})` };
  ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
  ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
  
  ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
  
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).numFmt = '0.0%';
  
  ws.mergeCells(row, 1, row, 3);
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.textbookLightBg),
      font: { ...FONT_BOLD, color: { argb: COLORS.textbookColor } },
      border: THIN_BORDER
    });
  }
  row++;
  
  // Spacing
  ws.getRow(row).height = 10;
  row++;
  
  const prepStartRow = row;
  
  // Prep Product Rows
  for (let i = 0; i < PREP_PRODUCTS.length; i++) {
    const prod = PREP_PRODUCTS[i];
    const data = custData[prod] || { qty: 0, rtn_qty: 0 };
    const bgColor = COLORS.dataBgOdd;
    
    ws.getCell(row, 1).value = idx++;
    ws.getCell(row, 2).value = '備課';
    ws.getCell(row, 3).value = getDisplayName(prod);
    ws.getCell(row, 4).value = data.qty;
    ws.getCell(row, 5).value = data.rtn_qty;
    ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
    ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
    
    ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 3).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
    
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).numFmt = '#,##0';
    ws.getCell(row, 7).numFmt = '0.0%';
    
    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER
      });
    }
    row++;
  }
  const prepEndRow = row - 1;
  const prepSubtotalRow = row;
  
  // Prep Subtotal Row
  ws.getCell(row, 1).value = '備課 小計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = { formula: `=SUM(D${prepStartRow}:D${prepEndRow})` };
  ws.getCell(row, 5).value = { formula: `=SUM(E${prepStartRow}:E${prepEndRow})` };
  ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
  ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
  
  ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
  
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).numFmt = '0.0%';
  
  ws.mergeCells(row, 1, row, 3);
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.prepLightBg),
      font: { ...FONT_BOLD, color: { argb: COLORS.prepColor } },
      border: THIN_BORDER
    });
  }
  row++;
  
  // Spacing
  ws.getRow(row).height = 10;
  row++;
  
  const grandTotalRow = row;
  
  // Grand Total Row (Row 16)
  ws.getCell(row, 1).value = '總合計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = { formula: `=D${textbookSubtotalRow}+D${prepSubtotalRow}` };
  ws.getCell(row, 5).value = { formula: `=E${textbookSubtotalRow}+E${prepSubtotalRow}` };
  ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
  ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
  
  ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
  
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).numFmt = '0.0%';
  
  ws.mergeCells(row, 1, row, 3);
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM
    });
  }
  
  // Conditional formatting for negative net
  ws.addConditionalFormatting({
    ref: `F${textbookStartRow}:F${grandTotalRow}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: {
          fill: solidFill(COLORS.negativeBg),
          font: { color: { argb: 'FF9C0006' } }
        }
      }
    ]
  });

  // Set Column Widths
  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 35;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 14;
  ws.getColumn(7).width = 12;
  
  ws.views = [{ showGridLines: true }];
}

function buildRegionSummarySheet(ws, zoneName, customersInZone, zoneProductTotals, customerMeta) {
  let row = 1;
  const COLS = 7;
  
  // Title Row
  ws.mergeCells(row, 1, row, 4);
  ws.getCell(row, 1).value = `金安出版社 高中台語出退貨統計表 — ${zoneName}總表`;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  
  ws.mergeCells(row, 5, row, COLS);
  ws.getCell(row, 5).value = '2025/04/26 ~ 2025/10/25';
  ws.getCell(row, 5).font = FONT_DATE;
  ws.getCell(row, 5).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
  }
  ws.getRow(row).height = 30;
  row++;
  
  // Spacing
  ws.getRow(row).height = 10;
  row++;
  
  // Section 1 Header
  ws.mergeCells(row, 1, row, COLS);
  ws.getCell(row, 1).value = '  一、產品銷售明細匯總';
  ws.getCell(row, 1).font = { ...FONT_BOLD, size: 11 };
  ws.getCell(row, 1).fill = solidFill(COLORS.subHeaderBg);
  ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getCell(row, 1).border = THIN_BORDER;
  for (let c = 2; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 24;
  row++;
  
  // Table 1 Headers
  const headers = ['序號', '產品類別', '產品名稱', '出貨數量', '退貨量', '實際出貨', '退書率'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: { horizontal: 'center', vertical: 'middle' }
    });
  }
  ws.getRow(row).height = 24;
  row++;
  
  const textbookStartRow = row;
  let idx = 1;
  
  // Textbook rows
  for (let i = 0; i < TEXTBOOK_PRODUCTS.length; i++) {
    const prod = TEXTBOOK_PRODUCTS[i];
    const data = zoneProductTotals[prod] || { qty: 0, rtn_qty: 0 };
    const bgColor = i % 2 === 0 ? COLORS.dataBgOdd : COLORS.dataBgEven;
    
    ws.getCell(row, 1).value = idx++;
    ws.getCell(row, 2).value = '課本';
    ws.getCell(row, 3).value = getDisplayName(prod);
    ws.getCell(row, 4).value = data.qty;
    ws.getCell(row, 5).value = data.rtn_qty;
    ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
    ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
    
    ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 3).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
    
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).numFmt = '#,##0';
    ws.getCell(row, 7).numFmt = '0.0%';
    
    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER
      });
    }
    row++;
  }
  const textbookEndRow = row - 1;
  const textbookSubtotalRow = row;
  
  // Textbook Subtotal
  ws.getCell(row, 1).value = '課本 小計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = { formula: `=SUM(D${textbookStartRow}:D${textbookEndRow})` };
  ws.getCell(row, 5).value = { formula: `=SUM(E${textbookStartRow}:E${textbookEndRow})` };
  ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
  ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
  
  ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
  
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).numFmt = '0.0%';
  
  ws.mergeCells(row, 1, row, 3);
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.textbookLightBg),
      font: { ...FONT_BOLD, color: { argb: COLORS.textbookColor } },
      border: THIN_BORDER
    });
  }
  row++;
  
  // Spacing row
  ws.getRow(row).height = 10;
  row++;
  
  const prepStartRow = row;
  // Prep book rows
  for (let i = 0; i < PREP_PRODUCTS.length; i++) {
    const prod = PREP_PRODUCTS[i];
    const data = zoneProductTotals[prod] || { qty: 0, rtn_qty: 0 };
    const bgColor = COLORS.dataBgOdd;
    
    ws.getCell(row, 1).value = idx++;
    ws.getCell(row, 2).value = '備課';
    ws.getCell(row, 3).value = getDisplayName(prod);
    ws.getCell(row, 4).value = data.qty;
    ws.getCell(row, 5).value = data.rtn_qty;
    ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
    ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
    
    ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 2).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 3).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
    
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).numFmt = '#,##0';
    ws.getCell(row, 7).numFmt = '0.0%';
    
    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER
      });
    }
    row++;
  }
  const prepEndRow = row - 1;
  const prepSubtotalRow = row;
  
  // Prep Subtotal
  ws.getCell(row, 1).value = '備課 小計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = { formula: `=SUM(D${prepStartRow}:D${prepEndRow})` };
  ws.getCell(row, 5).value = { formula: `=SUM(E${prepStartRow}:E${prepEndRow})` };
  ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
  ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
  
  ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
  
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).numFmt = '0.0%';
  
  ws.mergeCells(row, 1, row, 3);
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.prepLightBg),
      font: { ...FONT_BOLD, color: { argb: COLORS.prepColor } },
      border: THIN_BORDER
    });
  }
  row++;
  
  // Spacing row
  ws.getRow(row).height = 10;
  row++;
  
  // Grand Total row
  const productGrandTotalRow = row;
  ws.getCell(row, 1).value = '總合計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = { formula: `=D${textbookSubtotalRow}+D${prepSubtotalRow}` };
  ws.getCell(row, 5).value = { formula: `=E${textbookSubtotalRow}+E${prepSubtotalRow}` };
  ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
  ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
  
  ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
  
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).numFmt = '0.0%';
  
  ws.mergeCells(row, 1, row, 3);
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM
    });
  }
  row++;
  
  // Spacing rows
  ws.getRow(row).height = 15;
  row++;
  ws.getRow(row).height = 15;
  row++;
  
  // Section 2 Header
  ws.mergeCells(row, 1, row, COLS);
  ws.getCell(row, 1).value = '  二、經銷商(客戶)出退貨匯總';
  ws.getCell(row, 1).font = { ...FONT_BOLD, size: 11 };
  ws.getCell(row, 1).fill = solidFill(COLORS.subHeaderBg);
  ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getCell(row, 1).border = THIN_BORDER;
  for (let c = 2; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 24;
  row++;
  
  // Table 2 Headers
  const custHeaders = ['序號', '客戶名稱', '負責業務', '出貨數量', '退貨量', '實際出貨', '退書率'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = custHeaders[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: { horizontal: 'center', vertical: 'middle' }
    });
  }
  ws.getRow(row).height = 24;
  row++;
  
  const custStartRow = row;
  let custIdx = 1;
  
  const sortedCustomers = customersInZone.sort();
  
  for (const cust of sortedCustomers) {
    const meta = customerMeta[cust];
    const bgColor = custIdx % 2 === 0 ? COLORS.dataBgEven : COLORS.dataBgOdd;
    
    ws.getCell(row, 1).value = custIdx++;
    ws.getCell(row, 2).value = cust;
    ws.getCell(row, 3).value = meta.sales;
    ws.getCell(row, 4).value = { formula: `='${cust}'!D16` };
    ws.getCell(row, 5).value = { formula: `='${cust}'!E16` };
    ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
    ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
    
    ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 2).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell(row, 3).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
    
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).numFmt = '#,##0';
    ws.getCell(row, 7).numFmt = '0.0%';
    
    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bgColor),
        font: FONT_DATA,
        border: THIN_BORDER
      });
    }
    row++;
  }
  const custEndRow = row - 1;
  const custSummaryTotalRow = row;
  
  // Table 2 Summary Row
  ws.getCell(row, 1).value = '經銷商合計';
  ws.getCell(row, 2).value = '';
  ws.getCell(row, 3).value = '';
  ws.getCell(row, 4).value = { formula: `=SUM(D${custStartRow}:D${custEndRow})` };
  ws.getCell(row, 5).value = { formula: `=SUM(E${custStartRow}:E${custEndRow})` };
  ws.getCell(row, 6).value = { formula: `=D${row}-E${row}` };
  ws.getCell(row, 7).value = { formula: `=IF(D${row}>0, E${row}/D${row}, 0)` };
  
  ws.getCell(row, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(row, 4).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };
  ws.getCell(row, 7).alignment = { horizontal: 'right', vertical: 'middle' };
  
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).numFmt = '0.0%';
  
  ws.mergeCells(row, 1, row, 3);
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM
    });
  }
  
  // Conditional formatting for negative net
  ws.addConditionalFormatting({
    ref: `F${textbookStartRow}:F${productGrandTotalRow}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: {
          fill: solidFill(COLORS.negativeBg),
          font: { color: { argb: 'FF9C0006' } }
        }
      }
    ]
  });
  
  ws.addConditionalFormatting({
    ref: `F${custStartRow}:F${custSummaryTotalRow}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'lessThan',
        formulae: ['0'],
        style: {
          fill: solidFill(COLORS.negativeBg),
          font: { color: { argb: 'FF9C0006' } }
        }
      }
    ]
  });

  // Set widths
  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 25;
  ws.getColumn(3).width = 35;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 14;
  ws.getColumn(7).width = 12;
  
  ws.views = [{ showGridLines: true }];
}

// =================================================================
// Main Process
// =================================================================
async function main() {
  const queryResultPath = path.join(__dirname, 'query_results_hs_tw.json');
  if (!fs.existsSync(queryResultPath)) {
    console.error(`ERROR: Query results file not found at: ${queryResultPath}`);
    process.exit(1);
  }
  
  const records = JSON.parse(fs.readFileSync(queryResultPath, 'utf-8'));
  console.log(`Loaded ${records.length} records.`);
  
  // 1. Group records by customer, and sum quantities per product
  const customerData = {}; 
  const customerMeta = {}; 
  
  for (const r of records) {
    const cust = r.customer;
    if (!customerData[cust]) {
      customerData[cust] = {};
      customerMeta[cust] = { zone: r.zone, sales: r.sales };
    }
    if (!customerData[cust][r.product]) {
      customerData[cust][r.product] = { qty: 0, rtn_qty: 0 };
    }
    customerData[cust][r.product].qty += r.qty;
    customerData[cust][r.product].rtn_qty += r.rtn_qty;
  }
  
  // Filter for North region (北區) only
  const northCustomers = [];
  for (const cust of Object.keys(customerData)) {
    if (customerMeta[cust].zone === '北區') {
      northCustomers.push(cust);
    }
  }
  
  console.log(`Found ${northCustomers.length} customers in 北區.`);
  
  // Compute product totals for 北區
  const northProductTotals = {};
  for (const prod of ALL_CANONICAL_PRODUCTS) {
    northProductTotals[prod] = { qty: 0, rtn_qty: 0 };
  }
  
  for (const cust of northCustomers) {
    const data = customerData[cust];
    for (const prod of ALL_CANONICAL_PRODUCTS) {
      if (data[prod]) {
        northProductTotals[prod].qty += data[prod].qty;
        northProductTotals[prod].rtn_qty += data[prod].rtn_qty;
      }
    }
  }
  
  // Create Excel workbook
  const wb = new ExcelJS.Workbook();
  wb.creator = '金安出版社';
  wb.created = new Date();
  
  // Add 北區總表
  console.log("Creating 北區總表...");
  buildRegionSummarySheet(wb.addWorksheet('北區總表'), '北區', northCustomers, northProductTotals, customerMeta);
  
  // Add customer sheets for 北區
  const sortedCusts = northCustomers.sort();
  console.log(`Creating customer sheets for 北區 (${sortedCusts.length} customers)...`);
  for (const cust of sortedCusts) {
    const meta = customerMeta[cust];
    const data = customerData[cust];
    const sheetName = cust.length > 31 ? cust.substring(0, 31) : cust; // Excel limit
    const ws = wb.addWorksheet(sheetName);
    buildCustomerSheet(ws, cust, meta.zone, meta.sales, data);
  }
  
  // Save output
  const outputDir = path.join(__dirname, 'Output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  const outputPath = path.join(outputDir, '高中台語出書統計表_北區.xlsx');
  await wb.xlsx.writeFile(outputPath);
  
  // Put a copy in the root folder
  const rootOutputPath = path.join(__dirname, '高中台語出書統計表_北區.xlsx');
  await wb.xlsx.writeFile(rootOutputPath);
  
  console.log(`\n✓ Generated North Region Excel sheets: ${outputPath}`);
  console.log(`✓ Also copied to root: ${rootOutputPath}`);
}

main().catch(console.error);
