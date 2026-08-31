import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('raw_sales_page.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== Options for select[name='customer'] ===");
$("select[name='customer'] option").each((i, el) => {
  console.log(`Val: ${$(el).attr('value')} | Text: ${$(el).text().trim()}`);
});

console.log("\n=== Options for select[name='sales'] ===");
$("select[name='sales'] option").each((i, el) => {
  console.log(`Val: ${$(el).attr('value')} | Text: ${$(el).text().trim()}`);
});

console.log("\n=== Options for select[name='zone'] ===");
$("select[name='zone'] option").each((i, el) => {
  console.log(`Val: ${$(el).attr('value')} | Text: ${$(el).text().trim()}`);
});
