import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('raw_sales_page.html', 'utf-8');
const $ = cheerio.load(html);

// Print table headers
console.log("=== Headers ===");
$('table.list thead tr td').each((i, el) => {
  console.log(`Col ${i+1}: ${$(el).text().trim()} | name: ${$(el).attr('name') || ''}`);
});

console.log("\n=== First Data Row ===");
const firstRow = $('table.list tbody tr.data').first();
firstRow.find('td').each((i, el) => {
  console.log(`Col ${i+1}: ${$(el).text().trim()} | name: ${$(el).attr('name') || ''}`);
});

// Also print the full html of the first row
console.log("\n=== First Row HTML ===");
console.log(firstRow.html());
