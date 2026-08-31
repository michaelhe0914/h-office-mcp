import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('deliver_date_page.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== Deliver Date Page ===");
$('form').each((i, el) => {
  console.log(`Form ${i}: action=${$(el).attr('action')} method=${$(el).attr('method')}`);
  $(el).find('input, select').each((j, inp) => {
    console.log(`  Input: name=${$(inp).attr('name')} type=${$(inp).attr('type') || inp.name}`);
  });
});

console.log("\nHeaders:");
$('table tr').first().find('td, th').each((i, el) => {
  console.log(`  Col ${i+1}: ${$(el).text().trim()}`);
});

console.log("\nFirst 10 rows:");
$('table tr').slice(1, 10).each((i, row) => {
  const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
  console.log(`  Row ${i+1}: ${cells.join(' | ')}`);
});
