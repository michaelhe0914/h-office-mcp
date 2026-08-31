import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('raw_sales_page.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== Links in table ===");
$('table.list tbody tr.data a').each((i, el) => {
  console.log(`Text: ${$(el).text().trim()} | href: ${$(el).attr('href')}`);
});

console.log("\n=== Checking other forms or scripts ===");
$('script').each((i, el) => {
  const text = $(el).text();
  if (text.includes('click') || text.includes('href') || text.includes('window.location')) {
    console.log(`Script ${i}: ${text.substring(0, 300)}...`);
  }
});
