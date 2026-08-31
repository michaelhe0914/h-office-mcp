import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('dealer_deliver_date.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== Title ===");
console.log($('title').text());

console.log("\n=== Table Headers ===");
$('table tr').first().find('td, th').each((i, el) => {
  console.log(`Col ${i+1}: ${$(el).text().trim()}`);
});

console.log("\n=== First 10 rows ===");
$('table tr').slice(1, 20).each((i, row) => {
  const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
  if (cells.length > 0) {
    console.log(`Row ${i+1}: ${cells.join(' | ')}`);
  }
});
