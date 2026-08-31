import fs from 'fs';
import * as cheerio from 'cheerio';

const files = ['check_page.html', 'setting_page.html'];

for (const file of files) {
  if (!fs.existsSync(file)) continue;
  console.log(`\n===================================`);
  console.log(`File: ${file}`);
  console.log(`===================================`);
  
  const html = fs.readFileSync(file, 'utf-8');
  const $ = cheerio.load(html);
  
  // Print headers and links
  console.log("Links:");
  $('a').each((i, el) => {
    console.log(`  - Text: ${$(el).text().trim()} | href: ${$(el).attr('href')}`);
  });
  
  // Print first few forms or text
  console.log("Forms:");
  $('form').each((i, el) => {
    console.log(`  - Form: ${$(el).attr('action')} | method: ${$(el).attr('method')}`);
  });
}
