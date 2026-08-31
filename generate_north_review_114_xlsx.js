import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_all_regions_review.json'), 'utf-8'));

// Target classes
const PRODUCT_CLASSES = [
  "國中講義:複習講義(不含5-6)",
  "國中講義:新思維(不含5-6)",
  "國中講義:新思維5-6"
];

const CLASS_DISPLAY_NAMES = {
  "國中講義:複習講義(不含5-6)": "複習講義",
  "國中講義:新思維(不含5-6)": "新思維講義(不含5-6)",
  "國中講義:新思維5-6": "新思維講義(5-6)"
};

const COLORS = {
  titleBg:      'FF1F4E79',  // 深藍
  titleFont:    'FFFFFFFF',  // 白
  headerBg:     'FF4472C4',  // 中藍
  headerFont:   'FFFFFFFF',  // 白
  
  "國中講義:複習講義(不含5-6)": { bg: 'FFE2EFDA', text: 'FF375623' },      // 淡綠 / 深綠字
  "國中講義:新思維(不含5-6)": { bg: 'FFFCE4D6', text: 'FFC65911' },        // 淡橘 / 深橘字
  "國中講義:新思維5-6": { bg: 'FFE8D5F5', text: 'FF7030A0' },            // 淡紫 / 紫字
  
  totalBg:      'FFFFF2CC',  // 總合計淡黃
  totalFont:    'FF833C0B',  // 總合計褐字
  dataBgEven:   'FFF2F2F2',  // 偶數行灰底
  dataBgOdd:    'FFFFFFFF',  // 奇數行白底
  borderColor:  'FF000000',
};

// Font styles
const FONT_TITLE   = { name: '微軟正黑體', size: 14, bold: true, color: { argb: COLORS.titleFont } };
const FONT_HEADER  = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.headerFont } };
const FONT_DATA    = { name: '微軟正黑體', size: 10 };
const FONT_GRAND   = { name: '微軟正黑體', size: 11, bold: true, color: { argb: COLORS.totalFont } };

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

const productsByClass = {};
// Filter for 北區 and collect products
const northRecords = allRecords.filter(r => r.zone === '北區');

for (const r of northRecords) {
  const cleanName = r.product.replace(/^\*/, '');
  if (!productsByClass[r.productClass]) {
    productsByClass[r.productClass] = new Set();
  }
  productsByClass[r.productClass].add(cleanName);
}

for (const pc of PRODUCT_CLASSES) {
  if (productsByClass[pc]) {
    // Basic alphabetical sort for now
    productsByClass[pc] = [...productsByClass[pc]].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  }
}

function buildSheet(ws, sheetTitle, dataMap) {
  let row = 1;
  const COLS = 6; // 序號, 產品類別, 產品名稱, 出貨量, 退貨量, 淨量

  // ---- Title Row ----
  ws.mergeCells(row, 1, row, 4);
  ws.getCell(row, 1).value = `114上 北區複習與新思維講義出貨統計 — ${sheetTitle}`;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 5, row, COLS);
  ws.getCell(row, 5).value = '114.4.26~114.9.25';
  ws.getCell(row, 5).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 5).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 5).alignment = { horizontal: 'right', vertical: 'middle' };

  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
    if (c > 1 && c !== 5) {
      ws.getCell(row, c).fill = solidFill(COLORS.titleBg);
    }
  }
  ws.getRow(row).height = 28;
  row++;

  // ---- Header Row ----
  const headers = ['序號', '產品類別', '產品', '出貨量', '退貨量', '淨量'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER,
    });
  }
  ws.getRow(row).height = 20;
  row++;

  let idx = 1;
  const subtotalRows = [];

  for (const pc of PRODUCT_CLASSES) {
    const pcProducts = productsByClass[pc] || [];
    if (pcProducts.length === 0) continue;

    const classStartRow = row;
    const classColors = COLORS[pc] || { bg: 'FFFFFFFF', text: 'FF000000' };

    for (let pi = 0; pi < pcProducts.length; pi++) {
      const pName = pcProducts[pi];
      const record = (dataMap[pc] && dataMap[pc][pName]) || { qty: 0, rtn_qty: 0 };
      
      const isEven = idx % 2 === 0;
      const rowBgColor = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

      ws.getCell(row, 1).value = idx;
      ws.getCell(row, 2).value = CLASS_DISPLAY_NAMES[pc];
      ws.getCell(row, 3).value = pName; 

      ws.getCell(row, 4).value = record.qty || null;
      if (record.qty) ws.getCell(row, 4).numFmt = '#,##0';

      ws.getCell(row, 5).value = record.rtn_qty || null;
      if (record.rtn_qty) ws.getCell(row, 5).numFmt = '#,##0';

      ws.getCell(row, 6).value = { formula: `IF(OR(D${row}>0, E${row}>0), D${row}-E${row}, "")` };
      ws.getCell(row, 6).numFmt = '#,##0';

      for (let c = 1; c <= COLS; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(rowBgColor),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c === 3 ? LEFT : CENTER,
        });
      }

      ws.getCell(row, 2).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: classColors.text } };

      ws.getRow(row).height = 18;
      idx++;
      row++;
    }

    const classEndRow = row - 1;

    // Subtotal
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = `${CLASS_DISPLAY_NAMES[pc]} 小計`;
    ws.getCell(row, 3).value = '';
    ws.getCell(row, 4).value = { formula: `SUM(D${classStartRow}:D${classEndRow})` };
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = { formula: `SUM(E${classStartRow}:E${classEndRow})` };
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
    ws.getCell(row, 6).numFmt = '#,##0';

    const subtotalFont = { name: '微軟正黑體', size: 10, bold: true, color: { argb: classColors.text } };
    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(classColors.bg),
        font: subtotalFont,
        border: MEDIUM_BORDER_BOTTOM,
        alignment: c === 2 ? LEFT : CENTER,
      });
    }
    ws.getRow(row).height = 20;
    subtotalRows.push(row);
    row++;
  }

  // Spacing
  ws.getRow(row).height = 12;
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, { border: THIN_BORDER });
  }
  row++;

  // Grand Total
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '總合計';
  ws.getCell(row, 3).value = '';

  const orderFormula = subtotalRows.length > 0 ? subtotalRows.map(rNum => `D${rNum}`).join('+') : '0';
  ws.getCell(row, 4).value = { formula: orderFormula };
  ws.getCell(row, 4).numFmt = '#,##0';

  const returnFormula = subtotalRows.length > 0 ? subtotalRows.map(rNum => `E${rNum}`).join('+') : '0';
  ws.getCell(row, 5).value = { formula: returnFormula };
  ws.getCell(row, 5).numFmt = '#,##0';

  ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
  ws.getCell(row, 6).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 2 ? LEFT : CENTER,
    });
  }
  ws.getRow(row).height = 22;

  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 30;
  ws.getColumn(4).width = 12;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 12;

  ws.views = [{ showGridLines: true }];
}

async function generateWorkbook(filename, sheetTitle, customerDataMap, overallDataMap) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'System';
  wb.created = new Date();

  const wsSummary = wb.addWorksheet('總表總量', {
    pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });
  buildSheet(wsSummary, sheetTitle, overallDataMap);

  const sortedCustomers = Object.keys(customerDataMap).sort();
  for (const customer of sortedCustomers) {
    let sheetName = customer.replace(/[\\\/?*\[\]]/g, '');
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

    const ws = wb.addWorksheet(sheetName, {
      pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    
    buildSheet(ws, customer, customerDataMap[customer]);
  }

  const outputDir = path.join(__dirname, 'Output3');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = path.join(outputDir, filename);
  await wb.xlsx.writeFile(outputPath);
  console.log(`✓ Generated: ${outputPath} (${wb.worksheets.length} sheets)`);
}

async function main() {
  const salesRepData = {};
  
  for (const r of northRecords) {
    const cleanName = r.product.replace(/^\*/, '');
    const rep = r.sales || '未指定業務';

    if (!salesRepData[rep]) {
      salesRepData[rep] = { customerData: {}, summaryData: {} };
    }

    const repData = salesRepData[rep];
    if (!repData.customerData[r.customer]) repData.customerData[r.customer] = {};
    if (!repData.customerData[r.customer][r.productClass]) repData.customerData[r.customer][r.productClass] = {};
    if (!repData.customerData[r.customer][r.productClass][cleanName]) {
      repData.customerData[r.customer][r.productClass][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    repData.customerData[r.customer][r.productClass][cleanName].qty += r.qty;
    repData.customerData[r.customer][r.productClass][cleanName].rtn_qty += r.rtn_qty;

    if (!repData.summaryData[r.productClass]) repData.summaryData[r.productClass] = {};
    if (!repData.summaryData[r.productClass][cleanName]) {
      repData.summaryData[r.productClass][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    repData.summaryData[r.productClass][cleanName].qty += r.qty;
    repData.summaryData[r.productClass][cleanName].rtn_qty += r.rtn_qty;
  }

  console.log("\nGenerating per-sales-rep files in Output3...");
  const sortedReps = Object.keys(salesRepData).sort();
  for (const rep of sortedReps) {
    const filename = `${rep}_114上_北區複習與新思維講義統計(含5-6).xlsx`;
    const title = `${rep} 負責客戶小計`;
    await generateWorkbook(filename, title, salesRepData[rep].customerData, salesRepData[rep].summaryData);
  }

  console.log('\nAll files successfully generated in Output3!');
}

main().catch(console.error);
