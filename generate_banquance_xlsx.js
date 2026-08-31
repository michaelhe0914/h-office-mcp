import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// 1. Ensure directory '半全冊' exists
const outputDir = path.join(__dirname, '半全冊');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
  console.log(`Created folder: ${outputDir}`);
}

// 2. Load query records
const rawData = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_north_banquance.json'), 'utf-8'));
const northRecords = rawData.filter(r => r.zone === '北區');

console.log(`Loaded ${northRecords.length} North records for 國中考卷:複習卷-半全冊`);

// Styling constants
const COLORS = {
  titleBg:      'FF1F4E79',  // 深藍
  titleFont:    'FFFFFFFF',  // 白
  headerBg:     'FF4472C4',  // 中藍
  headerFont:   'FFFFFFFF',  // 白
  repHeaderBg:  'FF2E75B6',  // 業務標題底色
  repHeaderFont:'FFFFFFFF',  // 業務標題白字
  subHeaderBg:  'FFD6E4F0',  // 副標底色
  subHeaderFont:'FF1F4E79',  // 副標深藍字
  dataBgEven:   'FFF2F2F2',  // 偶數列淡灰底
  dataBgOdd:    'FFFFFFFF',  // 奇數列白底
  totalBg:      'FFFFF2CC',  // 總合計淡黃底
  totalFont:    'FF833C0B',  // 總合計棕字
  negativeBg:   'FFFFC7CE',  // 負數淡粉紅
  negativeFont: 'FFCC0000',  // 負數紅字
  borderColor:  'FF000000'
};

const FONT_TITLE   = { name: '微軟正黑體', size: 14, bold: true, color: { argb: COLORS.titleFont } };
const FONT_HEADER  = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.headerFont } };
const FONT_REP_HDR = { name: '微軟正黑體', size: 11, bold: true, color: { argb: COLORS.repHeaderFont } };
const FONT_SUBHDR  = { name: '微軟正黑體', size: 10, bold: true, color: { argb: COLORS.subHeaderFont } };
const FONT_DATA    = { name: '微軟正黑體', size: 10 };
const FONT_DATA_BOLD = { name: '微軟正黑體', size: 10, bold: true };
const FONT_TOTAL   = { name: '微軟正黑體', size: 11, bold: true, color: { argb: COLORS.totalFont } };

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
const LEFT   = { horizontal: 'left',   vertical: 'middle' };
const RIGHT  = { horizontal: 'right',  vertical: 'middle' };

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

// Organize data by sales rep
const salesRepsData = {};
for (const r of northRecords) {
  const rep = r.sales;
  if (!salesRepsData[rep]) {
    salesRepsData[rep] = {
      customers: {},
      records: []
    };
  }
  const repObj = salesRepsData[rep];
  repObj.records.push(r);

  if (!repObj.customers[r.customer]) {
    repObj.customers[r.customer] = [];
  }
  repObj.customers[r.customer].push(r);
}

const targetReps = ['何光傑', '李敏豪', '林智偉', '康晉瑋', '朱鵬學'];

// =================================================================
// Sheet 1: 北區業務與客戶統計
// =================================================================
function buildSalesSummarySheet(ws) {
  let row = 1;
  const COLS = 7;

  // Title
  ws.mergeCells(row, 1, row, 5);
  ws.getCell(row, 1).value = '2025.7.01~9.25 北區半全冊卷 — 業務與所屬客戶數量統計表';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 6, row, COLS);
  ws.getCell(row, 6).value = '統計區間：2025.7.01 ~ 2025.9.25';
  ws.getCell(row, 6).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 6).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 6).alignment = RIGHT;

  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 28;
  row++;

  // Subtitle info
  ws.mergeCells(row, 1, row, COLS);
  ws.getCell(row, 1).value = '【業務別總覽】北區各業務人員、所屬客戶數量及半全冊卷進退貨統計';
  ws.getCell(row, 1).font = FONT_SUBHDR;
  ws.getCell(row, 1).fill = solidFill(COLORS.subHeaderBg);
  ws.getCell(row, 1).alignment = LEFT;
  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 22;
  row++;

  // Table Headers
  const headers = ['項次', '業務姓名', '所屬客戶數量', '客戶名稱列表', '出貨總量 (張)', '退貨總量 (張)', '淨出貨量 (張)'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER
    });
  }
  ws.getRow(row).height = 22;
  row++;

  const startRow = row;
  let idx = 1;

  for (const rep of targetReps) {
    const repObj = salesRepsData[rep] || { customers: {}, records: [] };
    const custList = Object.keys(repObj.customers);
    const custCount = custList.length;
    const custStr = custList.join(', ');

    const isEven = idx % 2 === 0;
    const bg = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

    ws.getCell(row, 1).value = idx;
    ws.getCell(row, 2).value = rep;
    ws.getCell(row, 3).value = custCount;
    ws.getCell(row, 4).value = custStr;

    // Formulas for sums from detail records
    // We will put raw values or formulas summing the rep's customer details
    let repQty = 0, repRtn = 0;
    for (const r of repObj.records) {
      repQty += r.qty;
      repRtn += r.rtn_qty;
    }

    ws.getCell(row, 5).value = repQty;
    ws.getCell(row, 5).numFmt = '#,##0';

    ws.getCell(row, 6).value = repRtn;
    ws.getCell(row, 6).numFmt = '#,##0';

    ws.getCell(row, 7).value = { formula: `E${row}-F${row}` };
    ws.getCell(row, 7).numFmt = '#,##0';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bg),
        font: c === 2 || c === 3 ? FONT_DATA_BOLD : FONT_DATA,
        border: THIN_BORDER,
        alignment: c === 1 || c === 3 ? CENTER : (c === 2 ? CENTER : (c === 4 ? LEFT : RIGHT))
      });
    }

    ws.getRow(row).height = 20;
    idx++;
    row++;
  }

  const endRow = row - 1;

  // Grand Total Row
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '北區合計';
  ws.getCell(row, 3).value = { formula: `SUM(C${startRow}:C${endRow})` };
  ws.getCell(row, 4).value = `全北區共 5 位業務 / 29 家客戶`;
  ws.getCell(row, 5).value = { formula: `SUM(E${startRow}:E${endRow})` };
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = { formula: `SUM(F${startRow}:F${endRow})` };
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).value = { formula: `E${row}-F${row}` };
  ws.getCell(row, 7).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 3 ? CENTER : (c <= 2 ? CENTER : (c === 4 ? LEFT : RIGHT))
    });
  }
  ws.getRow(row).height = 24;
  row += 2;

  // Section 2: Detailed Customer Breakdown by Sales Rep Table
  ws.mergeCells(row, 1, row, COLS);
  ws.getCell(row, 1).value = '【客戶明細彙總表】北區各業務人員之個別客戶出貨統計明細';
  ws.getCell(row, 1).font = FONT_SUBHDR;
  ws.getCell(row, 1).fill = solidFill(COLORS.subHeaderBg);
  ws.getCell(row, 1).alignment = LEFT;
  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 22;
  row++;

  const headers2 = ['項次', '業務姓名', '客戶名稱', '出貨品項數', '出貨量 (張)', '退貨量 (張)', '淨出貨量 (張)'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers2[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER
    });
  }
  ws.getRow(row).height = 22;
  row++;

  let detailIdx = 1;
  const custStartRow = row;

  for (const rep of targetReps) {
    const repObj = salesRepsData[rep] || { customers: {} };
    const custEntries = Object.entries(repObj.customers);

    for (const [custName, recs] of custEntries) {
      let cQty = 0, cRtn = 0;
      for (const r of recs) {
        cQty += r.qty;
        cRtn += r.rtn_qty;
      }
      const cNet = cQty - cRtn;
      const isEven = detailIdx % 2 === 0;
      const bg = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

      ws.getCell(row, 1).value = detailIdx;
      ws.getCell(row, 2).value = rep;
      ws.getCell(row, 3).value = custName;
      ws.getCell(row, 4).value = recs.length;
      ws.getCell(row, 5).value = cQty;
      ws.getCell(row, 5).numFmt = '#,##0';
      ws.getCell(row, 6).value = cRtn;
      ws.getCell(row, 6).numFmt = '#,##0';
      ws.getCell(row, 7).value = { formula: `E${row}-F${row}` };
      ws.getCell(row, 7).numFmt = '#,##0';

      for (let c = 1; c <= COLS; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(bg),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c === 1 || c === 2 || c === 4 ? CENTER : (c === 3 ? LEFT : RIGHT)
        });
      }

      if (cNet < 0) {
        ws.getCell(row, 7).fill = solidFill(COLORS.negativeBg);
        ws.getCell(row, 7).font = { ...FONT_DATA, color: { argb: COLORS.negativeFont } };
      }

      ws.getRow(row).height = 20;
      detailIdx++;
      row++;
    }
  }

  const custEndRow = row - 1;

  // Customer Detail Total Row
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '客戶總計';
  ws.getCell(row, 3).value = `共 ${custEndRow - custStartRow + 1} 家客戶`;
  ws.getCell(row, 4).value = { formula: `SUM(D${custStartRow}:D${custEndRow})` };
  ws.getCell(row, 5).value = { formula: `SUM(E${custStartRow}:E${custEndRow})` };
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = { formula: `SUM(F${custStartRow}:F${custEndRow})` };
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).value = { formula: `E${row}-F${row}` };
  ws.getCell(row, 7).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 1 || c === 2 || c === 4 ? CENTER : (c === 3 ? LEFT : RIGHT)
    });
  }
  ws.getRow(row).height = 24;

  // Column Widths
  ws.getColumn(1).width = 8;   // 項次
  ws.getColumn(2).width = 14;  // 業務姓名
  ws.getColumn(3).width = 16;  // 所屬客戶數量 / 客戶名稱
  ws.getColumn(4).width = 45;  // 客戶名稱列表 / 出貨品項數
  ws.getColumn(5).width = 16;  // 出貨總量
  ws.getColumn(6).width = 16;  // 退貨總量
  ws.getColumn(7).width = 16;  // 淨出貨量

  ws.views = [{ showGridLines: true }];
}

// =================================================================
// Sheet 2: 半全冊卷全區產品總彙
// =================================================================
function buildProductSummarySheet(ws) {
  let row = 1;
  const COLS = 6;

  ws.mergeCells(row, 1, row, 4);
  ws.getCell(row, 1).value = '2025.7.01~9.25 北區半全冊卷 — 產品出貨統計總表';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 5, row, COLS);
  ws.getCell(row, 5).value = '統計區間：2025.7.01 ~ 2025.9.25';
  ws.getCell(row, 5).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 5).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 5).alignment = RIGHT;

  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 28;
  row++;

  const headers = ['項次', '產品類別', '產品名稱', '出貨總量 (張)', '退貨總量 (張)', '淨出貨量 (張)'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER
    });
  }
  ws.getRow(row).height = 22;
  row++;

  // Group products across North region
  const prodMap = {};
  for (const r of northRecords) {
    const pName = r.product.replace(/^\*/, '');
    if (!prodMap[pName]) {
      prodMap[pName] = { qty: 0, rtn: 0 };
    }
    prodMap[pName].qty += r.qty;
    prodMap[pName].rtn += r.rtn_qty;
  }

  const sortedProds = Object.keys(prodMap).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  const startRow = row;
  let idx = 1;

  for (const pName of sortedProds) {
    const data = prodMap[pName];
    const isEven = idx % 2 === 0;
    const bg = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

    ws.getCell(row, 1).value = idx;
    ws.getCell(row, 2).value = '國中考卷:複習卷-半全冊';
    ws.getCell(row, 3).value = pName;
    ws.getCell(row, 4).value = data.qty;
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = data.rtn;
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
    ws.getCell(row, 6).numFmt = '#,##0';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bg),
        font: FONT_DATA,
        border: THIN_BORDER,
        alignment: c === 1 ? CENTER : (c === 2 || c === 3 ? LEFT : RIGHT)
      });
    }

    ws.getRow(row).height = 20;
    idx++;
    row++;
  }

  const endRow = row - 1;

  // Grand Total
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '半全冊卷總計';
  ws.getCell(row, 3).value = `共 ${sortedProds.length} 項產品`;
  ws.getCell(row, 4).value = { formula: `SUM(D${startRow}:D${endRow})` };
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = { formula: `SUM(E${startRow}:E${endRow})` };
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
  ws.getCell(row, 6).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 1 ? CENTER : (c === 2 || c === 3 ? LEFT : RIGHT)
    });
  }
  ws.getRow(row).height = 24;

  ws.getColumn(1).width = 8;   // 項次
  ws.getColumn(2).width = 24;  // 產品類別
  ws.getColumn(3).width = 30;  // 產品名稱
  ws.getColumn(4).width = 16;  // 出貨總量
  ws.getColumn(5).width = 16;  // 退貨總量
  ws.getColumn(6).width = 16;  // 淨出貨量

  ws.views = [{ showGridLines: true }];
}

// =================================================================
// Sheet 3: 北區全區出貨總明細
// =================================================================
function buildFullDetailSheet(ws) {
  let row = 1;
  const COLS = 7;

  ws.mergeCells(row, 1, row, 5);
  ws.getCell(row, 1).value = '2025.7.01~9.25 北區半全冊卷 — 出貨明細全表';
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 6, row, COLS);
  ws.getCell(row, 6).value = '統計區間：2025.7.01 ~ 2025.9.25';
  ws.getCell(row, 6).font = { name: '微軟正黑體', size: 10, bold: true, color: { argb: 'FFFFFFCC' } };
  ws.getCell(row, 6).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 6).alignment = RIGHT;

  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 28;
  row++;

  const headers = ['項次', '業務姓名', '客戶名稱', '產品名稱', '出貨量 (張)', '退貨量 (張)', '淨出貨量 (張)'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = headers[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER
    });
  }
  ws.getRow(row).height = 22;
  row++;

  const startRow = row;
  let idx = 1;

  for (const rep of targetReps) {
    const repObj = salesRepsData[rep] || { customers: {} };
    const custEntries = Object.entries(repObj.customers);

    for (const [custName, recs] of custEntries) {
      for (const r of recs) {
        const cleanProduct = r.product.replace(/^\*/, '');
        const isEven = idx % 2 === 0;
        const bg = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

        ws.getCell(row, 1).value = idx;
        ws.getCell(row, 2).value = r.sales;
        ws.getCell(row, 3).value = r.customer;
        ws.getCell(row, 4).value = cleanProduct;
        ws.getCell(row, 5).value = r.qty;
        ws.getCell(row, 5).numFmt = '#,##0';
        ws.getCell(row, 6).value = r.rtn_qty;
        ws.getCell(row, 6).numFmt = '#,##0';
        ws.getCell(row, 7).value = { formula: `E${row}-F${row}` };
        ws.getCell(row, 7).numFmt = '#,##0';

        for (let c = 1; c <= COLS; c++) {
          styleCell(ws, row, c, {
            fill: solidFill(bg),
            font: FONT_DATA,
            border: THIN_BORDER,
            alignment: c === 1 || c === 2 ? CENTER : (c === 3 || c === 4 ? LEFT : RIGHT)
          });
        }

        if (r.qty - r.rtn_qty < 0) {
          ws.getCell(row, 7).fill = solidFill(COLORS.negativeBg);
          ws.getCell(row, 7).font = { ...FONT_DATA, color: { argb: COLORS.negativeFont } };
        }

        ws.getRow(row).height = 20;
        idx++;
        row++;
      }
    }
  }

  const endRow = row - 1;

  // Grand Total
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = '全區總計';
  ws.getCell(row, 3).value = `全北區共 ${endRow - startRow + 1} 筆明細`;
  ws.getCell(row, 4).value = '';
  ws.getCell(row, 5).value = { formula: `SUM(E${startRow}:E${endRow})` };
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = { formula: `SUM(F${startRow}:F${endRow})` };
  ws.getCell(row, 6).numFmt = '#,##0';
  ws.getCell(row, 7).value = { formula: `E${row}-F${row}` };
  ws.getCell(row, 7).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 1 || c === 2 ? CENTER : (c === 3 || c === 4 ? LEFT : RIGHT)
    });
  }
  ws.getRow(row).height = 24;

  ws.getColumn(1).width = 8;   // 項次
  ws.getColumn(2).width = 14;  // 業務
  ws.getColumn(3).width = 24;  // 客戶
  ws.getColumn(4).width = 30;  // 產品名稱
  ws.getColumn(5).width = 16;  // 出貨量
  ws.getColumn(6).width = 16;  // 退貨量
  ws.getColumn(7).width = 16;  // 淨出貨量

  ws.views = [{ showGridLines: true }];
}

// =================================================================
// Sheets 4 ~ 8: 個別業務專屬頁籤 / 檔案
// =================================================================
function buildRepSheet(ws, repName) {
  let row = 1;
  const COLS = 6;
  const repObj = salesRepsData[repName] || { customers: {}, records: [] };
  const custEntries = Object.entries(repObj.customers);

  // Title
  ws.mergeCells(row, 1, row, 4);
  ws.getCell(row, 1).value = `2025.7.01~9.25 業務：${repName} — 北區半全冊卷統計`;
  ws.getCell(row, 1).font = FONT_TITLE;
  ws.getCell(row, 1).fill = solidFill(COLORS.titleBg);
  ws.getCell(row, 1).alignment = LEFT;

  ws.mergeCells(row, 5, row, COLS);
  ws.getCell(row, 5).value = `客戶數：${custEntries.length} 家`;
  ws.getCell(row, 5).font = FONT_REP_HDR;
  ws.getCell(row, 5).fill = solidFill(COLORS.repHeaderBg);
  ws.getCell(row, 5).alignment = RIGHT;

  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 28;
  row++;

  // Customer Summary Table for this Rep
  ws.mergeCells(row, 1, row, COLS);
  ws.getCell(row, 1).value = `【所屬客戶統計彙總】${repName} 負責之 ${custEntries.length} 家客戶進退貨總量`;
  ws.getCell(row, 1).font = FONT_SUBHDR;
  ws.getCell(row, 1).fill = solidFill(COLORS.subHeaderBg);
  ws.getCell(row, 1).alignment = LEFT;
  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 22;
  row++;

  const custSummaryHeaders = ['項次', '客戶名稱', '出貨品項數', '出貨量 (張)', '退貨量 (張)', '淨出貨量 (張)'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = custSummaryHeaders[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER
    });
  }
  ws.getRow(row).height = 22;
  row++;

  const custSummaryStartRow = row;
  let cIdx = 1;

  for (const [custName, recs] of custEntries) {
    let cQty = 0, cRtn = 0;
    for (const r of recs) {
      cQty += r.qty;
      cRtn += r.rtn_qty;
    }
    const isEven = cIdx % 2 === 0;
    const bg = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

    ws.getCell(row, 1).value = cIdx;
    ws.getCell(row, 2).value = custName;
    ws.getCell(row, 3).value = recs.length;
    ws.getCell(row, 4).value = cQty;
    ws.getCell(row, 4).numFmt = '#,##0';
    ws.getCell(row, 5).value = cRtn;
    ws.getCell(row, 5).numFmt = '#,##0';
    ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
    ws.getCell(row, 6).numFmt = '#,##0';

    for (let c = 1; c <= COLS; c++) {
      styleCell(ws, row, c, {
        fill: solidFill(bg),
        font: FONT_DATA,
        border: THIN_BORDER,
        alignment: c === 1 || c === 3 ? CENTER : (c === 2 ? LEFT : RIGHT)
      });
    }

    if (cQty - cRtn < 0) {
      ws.getCell(row, 6).fill = solidFill(COLORS.negativeBg);
      ws.getCell(row, 6).font = { ...FONT_DATA, color: { argb: COLORS.negativeFont } };
    }

    ws.getRow(row).height = 20;
    cIdx++;
    row++;
  }

  const custSummaryEndRow = row - 1;

  // Rep Customer Subtotal
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = `${repName} 客戶小計`;
  ws.getCell(row, 3).value = { formula: `SUM(C${custSummaryStartRow}:C${custSummaryEndRow})` };
  ws.getCell(row, 4).value = { formula: `SUM(D${custSummaryStartRow}:D${custSummaryEndRow})` };
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = { formula: `SUM(E${custSummaryStartRow}:E${custSummaryEndRow})` };
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
  ws.getCell(row, 6).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 1 || c === 3 ? CENTER : (c === 2 ? LEFT : RIGHT)
    });
  }
  ws.getRow(row).height = 24;
  row += 2;

  // Section 2: Products Detail per Customer
  ws.mergeCells(row, 1, row, COLS);
  ws.getCell(row, 1).value = `【客戶出貨商品明細】${repName} 各客戶買進之半全冊卷品項明細`;
  ws.getCell(row, 1).font = FONT_SUBHDR;
  ws.getCell(row, 1).fill = solidFill(COLORS.subHeaderBg);
  ws.getCell(row, 1).alignment = LEFT;
  for (let c = 1; c <= COLS; c++) ws.getCell(row, c).border = THIN_BORDER;
  ws.getRow(row).height = 22;
  row++;

  const detailHeaders = ['項次', '客戶名稱', '產品名稱', '出貨量 (張)', '退貨量 (張)', '淨出貨量 (張)'];
  for (let c = 1; c <= COLS; c++) {
    ws.getCell(row, c).value = detailHeaders[c - 1];
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.headerBg),
      font: FONT_HEADER,
      border: THIN_BORDER,
      alignment: CENTER
    });
  }
  ws.getRow(row).height = 22;
  row++;

  const detailStartRow = row;
  let pIdx = 1;

  for (const [custName, recs] of custEntries) {
    for (const r of recs) {
      const cleanProduct = r.product.replace(/^\*/, '');
      const isEven = pIdx % 2 === 0;
      const bg = isEven ? COLORS.dataBgEven : COLORS.dataBgOdd;

      ws.getCell(row, 1).value = pIdx;
      ws.getCell(row, 2).value = custName;
      ws.getCell(row, 3).value = cleanProduct;
      ws.getCell(row, 4).value = r.qty;
      ws.getCell(row, 4).numFmt = '#,##0';
      ws.getCell(row, 5).value = r.rtn_qty;
      ws.getCell(row, 5).numFmt = '#,##0';
      ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
      ws.getCell(row, 6).numFmt = '#,##0';

      for (let c = 1; c <= COLS; c++) {
        styleCell(ws, row, c, {
          fill: solidFill(bg),
          font: FONT_DATA,
          border: THIN_BORDER,
          alignment: c === 1 ? CENTER : (c === 2 || c === 3 ? LEFT : RIGHT)
        });
      }

      if (r.qty - r.rtn_qty < 0) {
        ws.getCell(row, 6).fill = solidFill(COLORS.negativeBg);
        ws.getCell(row, 6).font = { ...FONT_DATA, color: { argb: COLORS.negativeFont } };
      }

      ws.getRow(row).height = 20;
      pIdx++;
      row++;
    }
  }

  const detailEndRow = row - 1;

  // Detail Total Row
  ws.getCell(row, 1).value = '';
  ws.getCell(row, 2).value = `${repName} 品項總計`;
  ws.getCell(row, 3).value = `共 ${detailEndRow - detailStartRow + 1} 筆產品計價項目`;
  ws.getCell(row, 4).value = { formula: `SUM(D${detailStartRow}:D${detailEndRow})` };
  ws.getCell(row, 4).numFmt = '#,##0';
  ws.getCell(row, 5).value = { formula: `SUM(E${detailStartRow}:E${detailEndRow})` };
  ws.getCell(row, 5).numFmt = '#,##0';
  ws.getCell(row, 6).value = { formula: `D${row}-E${row}` };
  ws.getCell(row, 6).numFmt = '#,##0';

  for (let c = 1; c <= COLS; c++) {
    styleCell(ws, row, c, {
      fill: solidFill(COLORS.totalBg),
      font: FONT_TOTAL,
      border: MEDIUM_BORDER_BOTTOM,
      alignment: c === 1 ? CENTER : (c === 2 || c === 3 ? LEFT : RIGHT)
    });
  }
  ws.getRow(row).height = 24;

  ws.getColumn(1).width = 8;   // 項次
  ws.getColumn(2).width = 24;  // 客戶名稱
  ws.getColumn(3).width = 30;  // 產品名稱
  ws.getColumn(4).width = 16;  // 出貨量
  ws.getColumn(5).width = 16;  // 退貨量
  ws.getColumn(6).width = 16;  // 淨出貨量

  ws.views = [{ showGridLines: true }];
}

// =================================================================
// Main Generation Script
// =================================================================
async function main() {
  console.log('Generating master XLSX workbook...');

  // 1. Create Master Workbook
  const masterWb = new ExcelJS.Workbook();
  masterWb.creator = '金安出版社';
  masterWb.created = new Date();

  // Sheet 1: 北區業務與客戶統計
  const ws1 = masterWb.addWorksheet('北區業務與客戶統計');
  buildSalesSummarySheet(ws1);

  // Sheet 2: 半全冊卷全區產品總彙
  const ws2 = masterWb.addWorksheet('半全冊卷全區產品總彙');
  buildProductSummarySheet(ws2);

  // Sheet 3: 北區全區出貨總明細
  const ws3 = masterWb.addWorksheet('北區全區出貨總明細');
  buildFullDetailSheet(ws3);

  // Sheets 4~8: Individual Sales Rep Sheets
  for (const rep of targetReps) {
    const wsRep = masterWb.addWorksheet(rep);
    buildRepSheet(wsRep, rep);
  }

  const masterPath = path.join(outputDir, '2025.7.01-9.25_北區半全冊卷出書統計表.xlsx');
  await masterWb.xlsx.writeFile(masterPath);
  console.log(`✓ Master XLSX generated: ${masterPath}`);

  // 2. Generate Individual Workbook per Sales Rep in '半全冊' folder
  for (const rep of targetReps) {
    const repWb = new ExcelJS.Workbook();
    repWb.creator = '金安出版社';
    repWb.created = new Date();

    const wsRep = repWb.addWorksheet(`${rep}_半全冊卷統計`);
    buildRepSheet(wsRep, rep);

    const repPath = path.join(outputDir, `${rep}_2025.7.01-9.25_北區半全冊卷統計.xlsx`);
    await repWb.xlsx.writeFile(repPath);
    console.log(`✓ Rep XLSX generated: ${repPath}`);
  }

  console.log('\nAll XLSX files successfully created in folder [半全冊]!');
}

main().catch(console.error);
