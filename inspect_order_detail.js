import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('order_detail.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== Title ===");
console.log($('title').text());

console.log("\n=== Page Content H3 ===");
$('h3').each((i, el) => console.log($(el).text().trim()));

console.log("\n=== Form fields / Info ===");
$('div#content').find('div, p, span').slice(0, 10).each((i, el) => {
  console.log($(el).text().trim());
});

console.log("\n=== Table Headers ===");
$('table tr').first().find('td, th').each((i, el) => {
  console.log(`Col ${i+1}: ${$(el).text().trim()}`);
});

console.log("\n=== Table Rows ===");
$('table tr').slice(1).each((i, row) => {
  const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
  if (cells.length > 0) {
    console.log(`Row ${i+1}: ${cells.join(' | ')}`);
  }
});
