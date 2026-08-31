import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('dealer_deliver_date.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== High Schools in dealer_deliver_date.html ===");
let count = 0;
$('table tr').slice(1).each((i, row) => {
  const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
  if (cells.length > 0) {
    const schoolName = cells[2] || "";
    if (schoolName.includes("高中") || schoolName.includes("中學") || schoolName.includes("高職") || schoolName.includes("高中職")) {
      console.log(`Row ${i+1}: ${cells.join(' | ')}`);
      count++;
    }
  }
});
console.log(`Total high schools found: ${count}`);
