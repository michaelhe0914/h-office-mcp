import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('battlefield_page.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== Title ===");
console.log($('title').text());

console.log("\n=== Form fields / Info ===");
$('form input, form select').each((i, el) => {
  console.log(`  - name: ${$(el).attr('name')} | type: ${$(el).attr('type') || el.name}`);
});

console.log("\n=== Table Headers ===");
$('table tr').first().find('td, th').each((i, el) => {
  console.log(`Col ${i+1}: ${$(el).text().trim()}`);
});

console.log("\n=== First 10 rows ===");
$('table tr').slice(1, 10).each((i, row) => {
  const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
  if (cells.length > 0) {
    console.log(`Row ${i+1}: ${cells.join(' | ')}`);
  }
});
