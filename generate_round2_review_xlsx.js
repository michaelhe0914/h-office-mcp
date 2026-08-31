import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_round2_review.json'), 'utf-8'));

// Target classes in order
const PRODUCT_CLASSES = [
  "國中講義:3900題",
  "國中講義:考前30天",
  "國中講義:UP+主題",
  "國中講義:歷屆試題"
];

// Display names for product classes
const CLASS_DISPLAY_NAMES = {
  "國中講義:3900題": "3900題",
  "國中講義:考前30天": "考前30天",
  "國中講義:UP+主題": "UP+主題",
  "國中講義:歷屆試題": "歷屆試題"
};

// Colors definition
const COLORS = {
  titleBg:      'FF1F4E79',  // 深藍
  titleFont:    'FFFFFFFF',  // 白
  headerBg:     'FF4472C4',  // 中藍
  headerFont:   'FFFFFFFF',  // 白
  
  "國中講義:3900題":  { bg: 'FFE2EFDA', text: 'FF375623' },  // 淡綠 / 深綠字
  "國中講義:考前30天": { bg: 'FFFCE4D6', text: 'FFC65911' },  // 淡橘 / 深橘字
  "國中講義:UP+主題":  { bg: 'FFE8D5F5', text: 'FF7030A0' },  // 淡紫 / 紫字
  "國中講義:歷屆試題": { bg: 'FFDAEEF3', text: 'FF1F4E79' },  // 淡藍 / 深藍字
  
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

// Subject ordering
const SUBJECT_ORDER = ['國文', '英語', '英文', '數學', '自然', '生物', '理化', '地科', '地球科學', '社會', '歷史', '地理', '公民'];

function getSubjectWeight(product) {
  for (let i = 0; i < SUBJECT_ORDER.length; i++) {
    if (product.includes(SUBJECT_ORDER[i])) return i;
  }
  return 99;
}

function compareProducts(p1, p2) {
  const s1 = getSubjectWeight(p1);
  const s2 = getSubjectWeight(p2);
  if (s1 !== s2) return s1 - s2;
  return p1.localeCompare(p2, 'zh-Hant');
}

// Clean product names & group by class
const productsByClass = {};
for (const r of allRecords) {
  const cleanName = r.product.replace(/^\*/, '');
  if (!productsByClass[r.productClass]) {
    productsByClass[r.productClass] = new Set();
  }
  productsByClass[r.productClass].add(cleanName);
}

for (const pc of PRODUCT_CLASSES) {
  if (productsByClass[pc]) {
    productsByClass[pc] = [...productsByClass[pc]].sort(compareProducts);
  }
}

// Build standard product detail sheet
function buildSheet(ws, sheetTitle, dataMap) {
  let row = 1;
  const COLS = 7; // 序號, 產品類別, 產品名稱, 訂量, 退量, 淨出貨, 退書率

  // ---- Title Row ----
  ws.mergeCells(row, 1, row, 5);
  ws.getCell(row, 1).value = `2025 第2波複習(3900,考前30,UP+,歷屆) — ${sheetTitle}`;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 6, row, COLS);
  ws.getCell(row, 6).value = '2025.7.01~2025.9.25';
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

      ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, "")` };
      ws.getCell(row, 7).numFmt = '0.0%';

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

    // ---- Subtotal Row ----
    ws.getCell(row, 1).value = '';
    ws.getCell(row, 2).value = `${CLASS_DISPLAY_NAMES[pc]} 小計`;
    ws.getCell(row, 3).value = '';

    ws.getCell(row, 4).value = { formula: `SUM(D${classStartRow}:D${classEndRow})` };
    ws.getCell(row, 4).numFmt = '#,##0';

    ws.getCell(row, 5).value = { formula: `SUM(E${classStartRow}:E${classEndRow})` };
    ws.getCell(row, 5).numFmt = '#,##0';

    ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
    ws.getCell(row, 6).numFmt = '#,##0';

    ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, 0)` };
    ws.getCell(row, 7).numFmt = '0.0%';

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

  // Spacing Row
  ws.getRow(row).height = 12;
  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, { border: THIN_BORDER });
  }
  row++;

  // Grand Total Row
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

  ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, 0)` };
  ws.getCell(row, 7).numFmt = '0.0%';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 2 ? LEFT : CENTER,
    });
  }
  ws.getRow(row).height = 22;

  // Set column widths
  ws.getColumn(1).width = 6;   // 序號
  ws.getColumn(2).width = 16;  // 產品類別
  ws.getColumn(3).width = 30;  // 產品名稱
  ws.getColumn(4).width = 12;  // 訂量
  ws.getColumn(5).width = 12;  // 退量
  ws.getColumn(6).width = 12;  // 淨出貨
  ws.getColumn(7).width = 12;  // 退書率

  ws.views = [{ showGridLines: true }];
}

// Build Sales Reps Overview Sheet (業務與所屬客戶統計)
function buildRepOverviewSheet(ws, salesRepStats) {
  let row = 1;
  const COLS = 7;

  ws.mergeCells(row, 1, row, 5);
  ws.getCell(row, 1).value = '北區各業務負責客戶數與出貨總量統計';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 6, row, COLS);
  ws.getCell(row, 6).value = '2025.7.01~2025.9.25';
  ws.getCell(row, 6).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 6).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 6).alignment = { horizontal: 'right', vertical: 'middle' };

  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).border = THIN_BORDER;
    if (c > 1 && c !== 6) ws.getCell(row, c).fill = solidFill(COLORS.titleBg);
  }
  ws.getRow(row).height = 28;
  row++;

  const headers = ['序號', '業務員', '所屬客戶數量', '總訂量', '總退量', '淨出貨', '退書率'];
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

  const startRow = row;
  const reps = Object.keys(salesRepStats)
    .filter(k => !k.startsWith('_'))
    .sort((a, b) => a.localeCompare(b, 'zh-Hant'));

  for (let i = 0; i < reps.length; i++) {
    const rep = reps[i];
    const stat = salesRepStats[rep];
    const isEven = (i + 1) % 2 === 0;
    const rowBgColor = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

    ws.getCell(row, 1).value = i + 1;
    ws.getCell(row, 2).value = rep;
    ws.getCell(row, 3).value = stat.customerCount;
    ws.getCell(row, 4).value = stat.qty;
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = stat.rtn_qty;
    ws.getCell(row, 5).numFmt = '#,##0';

    ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
    ws.getCell(row, 6).numFmt = '#,##0';

    ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, 0)` };
    ws.getCell(row, 7).numFmt = '0.0%';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(rowBgColor),
        font: FONT_DATA,
        border: THIN_BORDER,
        alignment: CENTER,
      });
    }
    ws.getRow(row).height = 20;
    row++;
  }

  const endRow = row - 1;

  // Grand Total Row
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '全北區合計';
  ws.getCell(row, 3).value = salesRepStats._overallCustomerCount;
  ws.getCell(row, 4).value = { formula: `SUM(D${startRow}:D${endRow})` };
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = { formula: `SUM(E${startRow}:E${endRow})` };
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).value = { formula: `IF(D${row}>0, E${row}/D${row}, 0)` };
  ws.getCell(row, 7).numFmt = '0.0%';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_GRAND,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: CENTER,
    });
  }
  ws.getRow(row).height = 22;

  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 16;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 14;
  ws.getColumn(6).width = 14;
  ws.getColumn(7).width = 14;

  ws.views = [{ showGridLines: true }];
}

// Generate Excel files into "第2波複習1" directory
async function main() {
  const targetDirName = '第2波複習1';
  const outputDir = path.join(__dirname, targetDirName);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
    console.log(`Created directory: ${outputDir}`);
  }

  // Aggregate Data
  const overallCustomerData = {};
  const overallSummaryData = {};
  const salesRepData = {};
  const overallNorthCustomers = new Set();

  for (const r of allRecords) {
    const cleanName = r.product.replace(/^\*/, '');
    const rep = r.sales || '未指定業務';
    const cust = r.customer;
    const pc = r.productClass;

    overallNorthCustomers.add(cust);

    // 1. Overall customer aggregation
    if (!overallCustomerData[cust]) overallCustomerData[cust] = {};
    if (!overallCustomerData[cust][pc]) overallCustomerData[cust][pc] = {};
    if (!overallCustomerData[cust][pc][cleanName]) {
      overallCustomerData[cust][pc][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    overallCustomerData[cust][pc][cleanName].qty += r.qty;
    overallCustomerData[cust][pc][cleanName].rtn_qty += r.rtn_qty;

    // 2. Overall product class aggregation
    if (!overallSummaryData[pc]) overallSummaryData[pc] = {};
    if (!overallSummaryData[pc][cleanName]) {
      overallSummaryData[pc][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    overallSummaryData[pc][cleanName].qty += r.qty;
    overallSummaryData[pc][cleanName].rtn_qty += r.rtn_qty;

    // 3. Sales rep aggregation
    if (!salesRepData[rep]) {
      salesRepData[rep] = {
        customerData: {},
        summaryData: {},
        customersSet: new Set(),
        totalQty: 0,
        totalRtn: 0
      };
    }
    salesRepData[rep].customersSet.add(cust);
    salesRepData[rep].totalQty += r.qty;
    salesRepData[rep].totalRtn += r.rtn_qty;

    const repData = salesRepData[rep];
    if (!repData.customerData[cust]) repData.customerData[cust] = {};
    if (!repData.customerData[cust][pc]) repData.customerData[cust][pc] = {};
    if (!repData.customerData[cust][pc][cleanName]) {
      repData.customerData[cust][pc][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    repData.customerData[cust][pc][cleanName].qty += r.qty;
    repData.customerData[cust][pc][cleanName].rtn_qty += r.rtn_qty;

    if (!repData.summaryData[pc]) repData.summaryData[pc] = {};
    if (!repData.summaryData[pc][cleanName]) {
      repData.summaryData[pc][cleanName] = { qty: 0, rtn_qty: 0 };
    }
    repData.summaryData[pc][cleanName].qty += r.qty;
    repData.summaryData[pc][cleanName].rtn_qty += r.rtn_qty;
  }

  // Construct sales rep statistics object for overview sheet
  const salesRepStats = {
    _overallCustomerCount: overallNorthCustomers.size
  };
  for (const [rep, d] of Object.entries(salesRepData)) {
    salesRepStats[rep] = {
      customerCount: d.customersSet.size,
      qty: d.totalQty,
      rtn_qty: d.totalRtn
    };
  }

  // --- Generate 1: Grand Overall North Region File ---
  console.log("\nGenerating grand overall file...");
  const wbOverall = new ExcelJS.Workbook();
  wbOverall.creator = '金安出版社';
  wbOverall.created = new Date();

  // Sheet 1: Sales Rep Overview
  const wsRepOverview = wbOverall.addWorksheet('業務與客戶統計', {
    pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });
  buildRepOverviewSheet(wsRepOverview, salesRepStats);

  // Sheet 2: Grand Total
  const wsOverallSummary = wbOverall.addWorksheet('總表總量', {
    pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });
  buildSheet(wsOverallSummary, '北區總量', overallSummaryData);

  // Per-customer sheets in overall file
  const sortedNorthCustomers = [...overallNorthCustomers].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  for (const cust of sortedNorthCustomers) {
    let sheetName = cust.replace(/[\\\/?*\[\]]/g, '');
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
    const ws = wbOverall.addWorksheet(sheetName, {
      pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    buildSheet(ws, cust, overallCustomerData[cust]);
  }

  const overallFilePath = path.join(outputDir, '北區_2025第2波複習出貨統計_總表.xlsx');
  await wbOverall.xlsx.writeFile(overallFilePath);
  console.log(`✓ Generated Grand Overall File: ${overallFilePath} (${wbOverall.worksheets.length} sheets)`);

  // --- Generate 2: Files per Sales Rep ---
  console.log("\nGenerating individual sales rep files...");
  const sortedReps = Object.keys(salesRepData).sort((a, b) => a.localeCompare(b, 'zh-Hant'));

  for (const rep of sortedReps) {
    const wbRep = new ExcelJS.Workbook();
    wbRep.creator = '金安出版社';
    wbRep.created = new Date();

    const repSummaryTitle = `${rep} 負責客戶小計`;

    // Summary Sheet
    const wsRepSummary = wbRep.addWorksheet('總量合計', {
      pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    buildSheet(wsRepSummary, repSummaryTitle, salesRepData[rep].summaryData);

    // Customer Sheets for this Rep
    const repCustomers = [...salesRepData[rep].customersSet].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
    for (const cust of repCustomers) {
      let sheetName = cust.replace(/[\\\/?*\[\]]/g, '');
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
      const ws = wbRep.addWorksheet(sheetName, {
        pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
      });
      buildSheet(ws, cust, salesRepData[rep].customerData[cust]);
    }

    const repFileName = `${rep}_2025第2波複習出貨統計.xlsx`;
    const repFilePath = path.join(outputDir, repFileName);
    await wbRep.xlsx.writeFile(repFilePath);
    console.log(`✓ Generated Rep File: ${repFilePath} (${wbRep.worksheets.length} sheets)`);
  }

  console.log(`\nAll files generated successfully in folder: ${outputDir}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
