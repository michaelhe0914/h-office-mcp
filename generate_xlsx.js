// Generate xlsx files for each sales rep in the format matching the reference template
// Each sales rep gets one xlsx with their customers' data as separate sheets

import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Load query results (includes all 5 reps including 康晉瑋)
const allRecords = JSON.parse(fs.readFileSync(path.join(__dirname, 'query_results_north_single_volume.json'), 'utf-8'));

// Target sales reps - 康晉瑋 is the correct name from the system
const targetSales = ["何光傑", "李敏豪", "林智偉", "康晉瑋", "朱鵬學"];

// =================================================================
// Product name parsing utilities
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

// =================================================================
// Build data structures
// =================================================================

function buildSalesData(records) {
  const salesData = {};
  
  for (const r of records) {
    const parsed = parseProduct(r.product, r.productClass);
    if (!parsed) continue;
    
    const { category, subject, version, grade } = parsed;
    const sales = r.sales;
    const customer = r.customer;
    const qty = r.qty;
    
    if (!salesData[sales]) salesData[sales] = {};
    if (!salesData[sales][customer]) salesData[sales][customer] = {};
    if (!salesData[sales][customer][category]) salesData[sales][customer][category] = {};
    if (!salesData[sales][customer][category][version]) salesData[sales][customer][category][version] = {};
    if (!salesData[sales][customer][category][version][subject]) salesData[sales][customer][category][version][subject] = {};
    
    const existing = salesData[sales][customer][category][version][subject][grade] || 0;
    salesData[sales][customer][category][version][subject][grade] = existing + qty;
  }
  
  return salesData;
}

// =================================================================
// Generate xlsx
// =================================================================

function generateSalesRepXlsx(salesRep, customerData) {
  const wb = XLSX.utils.book_new();
  const customers = Object.keys(customerData).sort();
  
  for (const customer of customers) {
    const data = customerData[customer];
    const sheetData = buildSheetData(customer, data);
    
    let sheetName = customer;
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
    sheetName = sheetName.replace(/[\\\/\?\*\[\]]/g, '');
    
    const ws = XLSX.utils.aoa_to_sheet(sheetData.rows);
    
    if (sheetData.merges.length > 0) {
      ws['!merges'] = sheetData.merges;
    }
    
    ws['!cols'] = sheetData.cols;
    
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }
  
  return wb;
}

function buildSheetData(customer, data) {
  const rows = [];
  const merges = [];
  let currentRow = 0;
  const grades = ['一', '三', '五'];
  
  // =============================================
  // Title Row
  // =============================================
  rows.push(['114年上期單冊講義', '', '', '', '', '', '日期：2025 製', '', '', '', '', '', '114.4.26~114.9.25']);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } });
  merges.push({ s: { r: 0, c: 6 }, e: { r: 0, c: 9 } });
  merges.push({ s: { r: 0, c: 12 }, e: { r: 0, c: 17 } });
  currentRow++;
  
  // =============================================
  // Section 1: 雙向 (國文/英語/數學/自然)
  // =============================================
  const sxData = data['雙向'] || {};
  const sxVersions = ['N', 'K', 'H'];
  const sxSubjects = ['國文', '英語', '數學', '自然'];
  
  // Header 1
  rows.push(['產品', '版本', customer, '國文', '', '', '英語', '', '', '數學', '', '', '自然', '', '', '訂書', '退貨', '實際', '總量']);
  merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow + 1, c: 0 } });
  merges.push({ s: { r: currentRow, c: 1 }, e: { r: currentRow + 1, c: 1 } });
  merges.push({ s: { r: currentRow, c: 2 }, e: { r: currentRow + 1, c: 2 } });
  merges.push({ s: { r: currentRow, c: 3 }, e: { r: currentRow, c: 5 } });
  merges.push({ s: { r: currentRow, c: 6 }, e: { r: currentRow, c: 8 } });
  merges.push({ s: { r: currentRow, c: 9 }, e: { r: currentRow, c: 11 } });
  merges.push({ s: { r: currentRow, c: 12 }, e: { r: currentRow, c: 14 } });
  currentRow++;
  
  // Header 2
  rows.push(['', '', '', '一', '三', '五', '一', '三', '五', '一', '三', '五', '一', '三', '五', '', '', '出貨', '']);
  currentRow++;
  
  // Data rows
  const sxDataStartRow = currentRow;
  for (let vi = 0; vi < sxVersions.length; vi++) {
    const v = sxVersions[vi];
    const vLabel = v === 'K' ? 'K版' : v === 'N' ? 'N版' : 'H版';
    const vData = sxData[v] || {};
    
    const row = [vi === 0 ? '雙向' : '', vLabel, '114上'];
    let rowTotal = 0;
    for (const subj of sxSubjects) {
      for (const grade of grades) {
        const qty = (vData[subj] && vData[subj][grade]) || 0;
        row.push(qty || '');
        rowTotal += qty;
      }
    }
    row.push(rowTotal || 0);
    row.push('');
    row.push(rowTotal || 0);
    row.push(rowTotal || 0);
    rows.push(row);
    currentRow++;
  }
  
  merges.push({ s: { r: sxDataStartRow, c: 0 }, e: { r: sxDataStartRow + sxVersions.length, c: 0 } });
  
  // Empty row
  rows.push([]);
  currentRow++;
  
  // 合計 row
  const sxTotals = ['', '合計', '114上'];
  let sxGrandTotal = 0;
  for (const subj of sxSubjects) {
    for (const grade of grades) {
      let total = 0;
      for (const v of sxVersions) {
        total += ((sxData[v] || {})[subj] || {})[grade] || 0;
      }
      sxTotals.push(total || 0);
      sxGrandTotal += total;
    }
  }
  sxTotals.push(sxGrandTotal || 0);
  sxTotals.push(0);
  sxTotals.push(sxGrandTotal || 0);
  sxTotals.push(sxGrandTotal || 0);
  rows.push(sxTotals);
  currentRow++;
  
  // Separator
  rows.push([]);
  rows.push([]);
  currentRow += 2;
  
  // =============================================
  // Section 2: 735 (國文/英語/數學/自然/地理/歷史)
  // =============================================
  const q735Data = data['735'] || {};
  const q735Versions = ['N', 'K', 'H'];
  const q735Subjects = ['國文', '英語', '數學', '自然', '地理', '歷史'];
  
  rows.push(['產品', '版本', customer, '國文', '', '', '英語', '', '', '數學', '', '', '自然', '', '', '地理', '', '', '歷史', '', '', '訂書', '退貨', '實際']);
  merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow + 1, c: 0 } });
  merges.push({ s: { r: currentRow, c: 1 }, e: { r: currentRow + 1, c: 1 } });
  merges.push({ s: { r: currentRow, c: 2 }, e: { r: currentRow + 1, c: 2 } });
  merges.push({ s: { r: currentRow, c: 3 }, e: { r: currentRow, c: 5 } });
  merges.push({ s: { r: currentRow, c: 6 }, e: { r: currentRow, c: 8 } });
  merges.push({ s: { r: currentRow, c: 9 }, e: { r: currentRow, c: 11 } });
  merges.push({ s: { r: currentRow, c: 12 }, e: { r: currentRow, c: 14 } });
  merges.push({ s: { r: currentRow, c: 15 }, e: { r: currentRow, c: 17 } });
  merges.push({ s: { r: currentRow, c: 18 }, e: { r: currentRow, c: 20 } });
  currentRow++;
  
  rows.push(['', '', '', '一', '三', '五', '一', '三', '五', '一', '三', '五', '一', '三', '五', '一', '三', '五', '一', '三', '五', '', '', '出貨']);
  currentRow++;
  
  const q735DataStartRow = currentRow;
  for (let vi = 0; vi < q735Versions.length; vi++) {
    const v = q735Versions[vi];
    const vLabel = v === 'K' ? 'K版' : v === 'N' ? 'N版' : 'H版';
    const vData = q735Data[v] || {};
    
    const row = [vi === 0 ? '735' : '', vLabel, '114上'];
    let rowTotal = 0;
    for (const subj of q735Subjects) {
      for (const grade of grades) {
        const qty = (vData[subj] && vData[subj][grade]) || 0;
        row.push(qty || '');
        rowTotal += qty;
      }
    }
    row.push(rowTotal || 0);
    row.push('');
    row.push(rowTotal || 0);
    rows.push(row);
    currentRow++;
  }
  
  merges.push({ s: { r: q735DataStartRow, c: 0 }, e: { r: q735DataStartRow + q735Versions.length, c: 0 } });
  
  rows.push([]);
  currentRow++;
  
  const q735Totals = ['', '合計', '114上'];
  let q735GrandTotal = 0;
  for (const subj of q735Subjects) {
    for (const grade of grades) {
      let total = 0;
      for (const v of q735Versions) {
        total += ((q735Data[v] || {})[subj] || {})[grade] || 0;
      }
      q735Totals.push(total || 0);
      q735GrandTotal += total;
    }
  }
  q735Totals.push(q735GrandTotal || 0);
  q735Totals.push(0);
  q735Totals.push(q735GrandTotal || 0);
  rows.push(q735Totals);
  currentRow++;
  
  rows.push([]);
  rows.push([]);
  currentRow += 2;
  
  // =============================================
  // Section 3: 試題篇 (英語/閱讀英文/數學)
  // =============================================
  const stpData = data['試題篇'] || {};
  const stpVersions = ['K', 'N', 'H'];
  const stpSubjects = ['英語', '閱讀英文', '數學'];
  
  rows.push(['產品', '版本', customer, '英語', '', '', '閱讀英文', '', '', '數學', '', '', '訂書', '退貨', '實際', '總量']);
  merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow + 1, c: 0 } });
  merges.push({ s: { r: currentRow, c: 1 }, e: { r: currentRow + 1, c: 1 } });
  merges.push({ s: { r: currentRow, c: 2 }, e: { r: currentRow + 1, c: 2 } });
  merges.push({ s: { r: currentRow, c: 3 }, e: { r: currentRow, c: 5 } });
  merges.push({ s: { r: currentRow, c: 6 }, e: { r: currentRow, c: 8 } });
  merges.push({ s: { r: currentRow, c: 9 }, e: { r: currentRow, c: 11 } });
  currentRow++;
  
  rows.push(['', '', '', '一', '三', '五', '一', '三', '五', '一', '三', '五', '', '', '出貨', '']);
  currentRow++;
  
  const stpDataStartRow = currentRow;
  for (let vi = 0; vi < stpVersions.length; vi++) {
    const v = stpVersions[vi];
    const vLabel = v === 'K' ? 'K版' : v === 'N' ? 'N版' : 'H版';
    const vData = stpData[v] || {};
    
    const row = [vi === 0 ? '試題篇' : '', vLabel, '114上'];
    let rowTotal = 0;
    for (const subj of stpSubjects) {
      for (const grade of grades) {
        const qty = (vData[subj] && vData[subj][grade]) || 0;
        row.push(qty || '');
        rowTotal += qty;
      }
    }
    row.push(rowTotal || 0);
    row.push('');
    row.push(rowTotal || 0);
    row.push(rowTotal || 0);
    rows.push(row);
    currentRow++;
  }
  
  merges.push({ s: { r: stpDataStartRow, c: 0 }, e: { r: stpDataStartRow + stpVersions.length, c: 0 } });
  
  rows.push([]);
  currentRow++;
  
  const stpTotals = ['', '合計', '114上'];
  let stpGrandTotal = 0;
  for (const subj of stpSubjects) {
    for (const grade of grades) {
      let total = 0;
      for (const v of stpVersions) {
        total += ((stpData[v] || {})[subj] || {})[grade] || 0;
      }
      stpTotals.push(total || 0);
      stpGrandTotal += total;
    }
  }
  stpTotals.push(stpGrandTotal || 0);
  stpTotals.push(0);
  stpTotals.push(stpGrandTotal || 0);
  stpTotals.push(stpGrandTotal || 0);
  rows.push(stpTotals);
  currentRow++;
  
  rows.push([]);
  rows.push([]);
  currentRow += 2;
  
  // =============================================
  // Section 4: 新講義 (數學/自然, 綜合版)
  // =============================================
  const xlData = data['新講義'] || {};
  const xlSubjects = ['數學', '自然'];
  
  rows.push(['產品', '版本', customer, '數學', '', '', '自然', '', '', '訂書', '退貨', '實際', '總量']);
  merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow + 1, c: 0 } });
  merges.push({ s: { r: currentRow, c: 1 }, e: { r: currentRow + 1, c: 1 } });
  merges.push({ s: { r: currentRow, c: 2 }, e: { r: currentRow + 1, c: 2 } });
  merges.push({ s: { r: currentRow, c: 3 }, e: { r: currentRow, c: 5 } });
  merges.push({ s: { r: currentRow, c: 6 }, e: { r: currentRow, c: 8 } });
  currentRow++;
  
  rows.push(['', '', '', '一', '三', '五', '一', '三', '五', '', '', '出貨', '']);
  currentRow++;
  
  const xlVData = xlData['綜合'] || {};
  const xlRow = ['新講義', '綜合版         合計', '114上'];
  let xlTotal = 0;
  for (const subj of xlSubjects) {
    for (const grade of grades) {
      const qty = (xlVData[subj] && xlVData[subj][grade]) || 0;
      xlRow.push(qty || '');
      xlTotal += qty;
    }
  }
  xlRow.push(xlTotal || 0);
  xlRow.push('');
  xlRow.push(xlTotal || 0);
  xlRow.push(xlTotal || 0);
  rows.push(xlRow);
  
  // Column widths
  const cols = [
    { wch: 10 }, { wch: 18 }, { wch: 8 },
    { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 7 }, { wch: 7 }, { wch: 7 },
  ];
  
  return { rows, merges, cols };
}

// =================================================================
// Main
// =================================================================

const salesData = buildSalesData(allRecords);

const outputDir = path.join(__dirname, 'Output');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

for (const salesRep of targetSales) {
  const customerData = salesData[salesRep];
  
  if (!customerData || Object.keys(customerData).length === 0) {
    console.log(`${salesRep}: No data found, skipping.`);
    continue;
  }
  
  const wb = generateSalesRepXlsx(salesRep, customerData);
  const outputPath = path.join(outputDir, `${salesRep}_114上單冊講義.xlsx`);
  XLSX.writeFile(wb, outputPath);
  console.log(`Generated: ${outputPath}`);
  console.log(`  Customers: ${Object.keys(customerData).length} - ${Object.keys(customerData).join(', ')}`);
  
  // Calculate total qty for verification
  let totalQty = 0;
  for (const customer of Object.keys(customerData)) {
    for (const cat of Object.keys(customerData[customer])) {
      for (const ver of Object.keys(customerData[customer][cat])) {
        for (const subj of Object.keys(customerData[customer][cat][ver])) {
          for (const grade of Object.keys(customerData[customer][cat][ver][subj])) {
            totalQty += customerData[customer][cat][ver][subj][grade];
          }
        }
      }
    }
  }
  console.log(`  Total qty: ${totalQty}`);
}

console.log('\nDone!');
