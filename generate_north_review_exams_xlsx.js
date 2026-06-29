import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_north_review_exams.json'), 'utf-8'));

// Target classes in order
const PRODUCT_CLASSES = [
  "國中考卷:複習卷-A卷",
  "國中考卷:複習卷-B卷",
  "國中考卷:複習卷-其他",
  "國中考卷:複習卷-新思維",
  "國中考卷:複習卷-半全冊"
];

// Display names for product classes
const CLASS_DISPLAY_NAMES = {
  "國中考卷:複習卷-A卷": "複習卷-A卷",
  "國中考卷:複習卷-B卷": "複習卷-B卷",
  "國中考卷:複習卷-其他": "複習卷-其他",
  "國中考卷:複習卷-新思維": "複習卷-新思維",
  "國中考卷:複習卷-半全冊": "複習卷-半全冊"
};

// Colors definition
const COLORS = {
  titleBg:      'FF1F4E79',  // 深藍
  titleFont:    'FFFFFFFF',  // 白
  headerBg:     'FF4472C4',  // 中藍
  headerFont:   'FFFFFFFF',  // 白
  
  // Section-specific background colors for subtotals and accents
  "國中考卷:複習卷-A卷": { bg: 'FFE2EFDA', text: 'FF375623' },      // 淡綠 / 深綠字
  "國中考卷:複習卷-B卷": { bg: 'FFFCE4D6', text: 'FFC65911' },      // 淡橘 / 深橘字
  "國中考卷:複習卷-其他": { bg: 'FFFFF2CC', text: 'FF7F6000' },      // 淡黃 / 深黃字
  "國中考卷:複習卷-新思維": { bg: 'FFE8D5F5', text: 'FF7030A0' },    // 淡紫 / 紫字
  "國中考卷:複習卷-半全冊": { bg: 'FFDAEEF3', text: 'FF1F4E79' },    // 淡藍 / 深藍字
  
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
const FONT_BOLD    = { name: '微軟正黑體', size: 10, bold: true };
const FONT_TOTAL   = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.totalFont } };
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

// Helpers
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

// Custom product sorting logic
const SUBJECT_ORDER = ['國文', '英語', '數學', '生物', '理化', '地球科學', '歷史', '地理', '公民'];

function getSubjectWeight(product) {
  for (let i = 0; i < SUBJECT_ORDER.length; i++) {
    if (product.includes(SUBJECT_ORDER[i]) || (SUBJECT_ORDER[i] === '英語' && product.includes('英文'))) {
      return i;
    }
  }
  return 99; // Unknown subject goes to end
}

function getVolumeWeight(product) {
  if (product.includes('(1-2)')) return 1;
  if (product.includes('(1-4)')) return 2;
  if (product.includes('(1-6)')) return 3;
  if (product.includes('(3-4)')) return 4;
  if (product.includes('(3-6)')) return 5;
  if (product.includes('(5-6)')) return 6;
  if (product.includes('(半冊)') || product.includes('半冊')) return 7;
  if (product.includes('(全)') || product.includes('全冊') || product.includes(' 全冊')) return 8;
  if (product.includes('半.全冊')) return 9;
  return 99;
}

function compareProducts(p1, p2) {
  // 1. Compare subject
  const s1 = getSubjectWeight(p1);
  const s2 = getSubjectWeight(p2);
  if (s1 !== s2) return s1 - s2;

  // 2. Compare volume/range
  const v1 = getVolumeWeight(p1);
  const v2 = getVolumeWeight(p2);
  if (v1 !== v2) return v1 - v2;

  // 3. Alphabetical fallback
  return p1.localeCompare(p2, 'zh-Hant');
}

// Clean and aggregate products by removing leading *
const productsByClass = {};
for (const r of allRecords) {
  const cleanName = r.product.replace(/^\*/, '');
  if (!productsByClass[r.productClass]) {
    productsByClass[r.productClass] = new Set();
  }
  productsByClass[r.productClass].add(cleanName);
}

// Sort products in each class
const sortedProductsList = [];
for (const pc of PRODUCT_CLASSES) {
  const pSet = productsByClass[pc] || new Set();
  const sorted = [...pSet].sort(compareProducts);
  productsByClass[pc] = sorted; // Replace Set with sorted Array
  for (const pName of sorted) {
    sortedProductsList.push({ productClass: pc, productName: pName });
  }
}

// 3. Function to build a styled sheet
function buildSheet(ws, sheetTitle, dataMap) {
  let row = 1;
  const COLS = 7; // 序號, 產品類別, 產品名稱, 訂量, 退量, 淨出貨, 退書率

  // ---- Title Row ----
  ws.mergeCells(row, 1, row, 5);
  ws.getCell(row, 1).value = `114上 北區複習卷出貨統計 — ${sheetTitle}`;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 6, row, COLS);
  ws.getCell(row, 6).value = '114.4.26~114.9.25';
  ws.getCell(row, 6).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 6).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };

  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
    if (c > 1 && c !== 6) {
      ws.getCell(row, c).fill = solidFill(COLORS.titleBg);
    }
  }
  ws.getRow(row).height = 28;
  row++;

  // ---- Header Row ----
  const headers = ['序號', '產品類別', '產品名稱', '訂量', '退量', '淨出貨', '退書率'];
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

  // ---- Write Data Rows ----
  let idx = 1;
  const subtotalRows = []; // Keep track of subtotal row numbers

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
      ws.getCell(row, 3).value = pName; // clean name

      // Display quantities (null if 0 for cleaner layout and correct Excel formula evaluation)
      ws.getCell(row, 4).value = record.qty || null;
      if (record.qty) ws.getCell(row, 4).numFmt = '#,##0';

      ws.getCell(row, 5).value = record.rtn_qty || null;
      if (record.rtn_qty) ws.getCell(row, 5).numFmt = '#,##0';

      // Net Shipment: Formula =Order - Return, but display blank if both are empty/zero to avoid `#VALUE!` or cluttered zeros.
      ws.getCell(row, 6).value = { formula: `IF(OR(D${row}>0, E${row}>0), D${row}-E${row}, "")` };
      ws.getCell(row, 6).numFmt = '#,##0';

      // Return Rate: Formula =Return / Order, but display blank if Order is empty/zero.
      ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, "")` };
      ws.getCell(row, 7).numFmt = '0.0%';

      // Formatting
      for (let c = 1; c <= COLS; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(rowBgColor),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c === 3 ? LEFT : CENTER,
        });
      }

      // Format category label column to be bold and colored
      ws.getCell(row, 2).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: classColors.text } };

      ws.getRow(row).height = 18;
      idx++;
      row++;
    }

    const classEndRow = row - 1;

    // ---- Subtotal Row (小計) ----
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = `${CLASS_DISPLAY_NAMES[pc]} 小計`;
    ws.getCell(row, 3).value = '';

    // Subtotal Order: SUM of quantities
    ws.getCell(row, 4).value = { formula: `SUM(D${classStartRow}:D${classEndRow})` };
    ws.getCell(row, 4).numFmt = '#,##0';

    // Subtotal Return: SUM of returns
    ws.getCell(row, 5).value = { formula: `SUM(E${classStartRow}:E${classEndRow})` };
    ws.getCell(row, 5).numFmt = '#,##0';

    // Subtotal Net: Order - Return
    ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
    ws.getCell(row, 6).numFmt = '#,##0';

    // Subtotal Rate: Return / Order
    ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, 0)` };
    ws.getCell(row, 7).numFmt = '0.0%';

    // Styling subtotal row
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

  // ---- Spacing Row ----
  ws.getRow(row).height = 12;
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, { border: THIN_BORDER });
  }
  row++;

  // ---- Grand Total Row (總合計) ----
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '總合計';
  ws.getCell(row, 3).value = '';

  // Order Grand Total: SUM of subtotals
  const orderFormula = subtotalRows.map(rNum => `D${rNum}`).join('+');
  ws.getCell(row, 4).value = { formula: orderFormula };
  ws.getCell(row, 4).numFmt = '#,##0';

  // Return Grand Total: SUM of subtotals
  const returnFormula = subtotalRows.map(rNum => `E${rNum}`).join('+');
  ws.getCell(row, 5).value = { formula: returnFormula };
  ws.getCell(row, 5).numFmt = '#,##0';

  // Net Grand Total: Order - Return
  ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
  ws.getCell(row, 6).numFmt = '#,##0';

  // Return Rate: Return / Order
  ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, 0)` };
  ws.getCell(row, 7).numFmt = '0.0%';

  // Styling Grand Total
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 2 ? LEFT : CENTER,
    });
  }
  ws.getRow(row).height = 22;

  // ---- Set Column Widths ----
  ws.getColumn(1).width = 6;   // 序號
  ws.getColumn(2).width = 16;  // 產品類別
  ws.getColumn(3).width = 30;  // 產品名稱
  ws.getColumn(4).width = 12;  // 訂量
  ws.getColumn(5).width = 12;  // 退量
  ws.getColumn(6).width = 12;  // 淨出貨
  ws.getColumn(7).width = 12;  // 退書率

  // Enable gridlines explicitly
  ws.views = [{ showGridLines: true }];
}

// 4. Function to generate a workbook for a given set of data
async function generateWorkbook(filename, sheetTitle, customerDataMap, overallDataMap) {
  const wb = new ExcelJS.Workbook();
  wb.creator = '金安出版社';
  wb.created = new Date();

  // Create Summary Sheet
  const wsSummary = wb.addWorksheet('總表總量', {
    pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });
  buildSheet(wsSummary, sheetTitle, overallDataMap);

  // Get and sort customer names
  const sortedCustomers = Object.keys(customerDataMap).sort();
  
  for (const customer of sortedCustomers) {
    let sheetName = customer.replace(/[\\\/?*\[\]]/g, '');
    if (sheetName.length > 31) {
      sheetName = sheetName.substring(0, 31);
    }

    const ws = wb.addWorksheet(sheetName, {
      pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    
    buildSheet(ws, customer, customerDataMap[customer]);
  }

  // Create Output2 directory if not exists
  const output2Dir = path.join(__dirname, 'Output2');
  if (!fs.existsSync(output2Dir)) {
    fs.mkdirSync(output2Dir, { recursive: true });
  }

  const outputPath = path.join(output2Dir, filename);
  await wb.xlsx.writeFile(outputPath);
  console.log(`✓ Generated: ${outputPath} (${wb.worksheets.length} sheets)`);
}

// 5. Main script execution
async function main() {
  // Aggregate overall customer data
  const overallCustomerData = {};
  const overallSummaryData = {};

  // Group data by sales rep
  // salesRep -> customer -> productClass -> cleanProductName -> { qty, rtn_qty }
  const salesRepData = {};
  
  for (const r of allRecords) {
    const cleanName = r.product.replace(/^\*/, '');
    const rep = r.sales || '未指定業務';

    // 1. Overall Aggregation
    if (!overallCustomerData[r.customer]) overallCustomerData[r.customer] = {};
    if (!overallCustomerData[r.customer][r.productClass]) overallCustomerData[r.customer][r.productClass] = {};
    if (!overallCustomerData[r.customer][r.productClass][cleanName]) {
      overallCustomerData[r.customer][r.productClass][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    overallCustomerData[r.customer][r.productClass][cleanName].qty += r.qty;
    overallCustomerData[r.customer][r.productClass][cleanName].rtn_qty += r.rtn_qty;

    if (!overallSummaryData[r.productClass]) overallSummaryData[r.productClass] = {};
    if (!overallSummaryData[r.productClass][cleanName]) {
      overallSummaryData[r.productClass][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    overallSummaryData[r.productClass][cleanName].qty += r.qty;
    overallSummaryData[r.productClass][cleanName].rtn_qty += r.rtn_qty;

    // 2. Sales Rep Specific Aggregation
    if (!salesRepData[rep]) {
      salesRepData[rep] = {
        customerData: {},
        summaryData: {}
      };
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

  // Generate the grand overall file
  console.log("\nGenerating grand overall file...");
  await generateWorkbook('114上_北區複習卷統計.xlsx', '總表總量', overallCustomerData, overallSummaryData);

  // Generate files for each sales rep
  console.log("\nGenerating per-sales-rep files...");
  const sortedReps = Object.keys(salesRepData).sort();
  for (const rep of sortedReps) {
    const filename = `${rep}_114上_北區複習卷統計.xlsx`;
    const title = `${rep} 負責客戶小計`;
    await generateWorkbook(filename, title, salesRepData[rep].customerData, salesRepData[rep].summaryData);
  }

  console.log('\nAll files successfully generated!');
}

main().catch(console.error);
