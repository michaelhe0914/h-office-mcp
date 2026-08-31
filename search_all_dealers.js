import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('deliver_date_page.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== All Links in deliver_date_page.html ===");
$('a').each((i, el) => {
  console.log(`${$(el).text().trim()} | href: ${$(el).attr('href')}`);
});
