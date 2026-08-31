import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('dealer_deliver_date.html', 'utf-8');
const $ = cheerio.load(html);

const years = new Set();
$('table tr').slice(1).each((i, row) => {
  const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
  if (cells.length > 0) {
    const status = cells[3] || "";
    const matches = status.match(/\d{4}/g);
    if (matches) {
      matches.forEach(y => years.add(y));
    }
  }
});

console.log("=== Years found in dealer_deliver_date.html ===");
console.log([...years].sort());
