import fs from 'fs';
import * as cheerio from 'cheerio';

const files = [
  'statistic_page.html',
  'class_count_page.html',
  'class_count_all_page.html',
  'deliver_date_page.html'
];

for (const file of files) {
  if (!fs.existsSync(file)) continue;
  console.log(`\n===================================`);
  console.log(`File: ${file}`);
  console.log(`===================================`);
  
  const html = fs.readFileSync(file, 'utf-8');
  const $ = cheerio.load(html);
  
  // Print form fields and inputs
  console.log("Inputs / Selects:");
  $('form input, form select').each((i, el) => {
    const name = $(el).attr('name');
    const type = $(el).attr('type') || el.name;
    const value = $(el).val();
    if (name) {
      console.log(`  - name: ${name} | type: ${type} | value: ${value}`);
    }
  });
  
  // Print table headers
  console.log("Table Headers:");
  $('table thead tr td, table tr.header td').each((i, el) => {
    console.log(`  - Col ${i+1}: ${$(el).text().trim()}`);
  });
}
