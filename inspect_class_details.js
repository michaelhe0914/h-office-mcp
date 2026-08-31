import fs from 'fs';
import * as cheerio from 'cheerio';

const files = [
  'class_count_detail_page.html',
  'class_count_all_detail_page.html'
];

for (const file of files) {
  if (!fs.existsSync(file)) continue;
  console.log(`\n===================================`);
  console.log(`File: ${file}`);
  console.log(`===================================`);
  
  const html = fs.readFileSync(file, 'utf-8');
  const $ = cheerio.load(html);
  
  // Print Title
  console.log(`Title: ${$('title').text().trim()}`);
  
  // Print first few table headers
  console.log("Table Headers:");
  $('table tr').first().find('td, th').each((i, el) => {
    console.log(`  - Col ${i+1}: ${$(el).text().trim()}`);
  });
  
  // Print first 5 rows
  console.log("First 5 rows:");
  $('table tr').slice(1, 6).each((i, row) => {
    const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
    console.log(`  - Row ${i+1}: ${cells.join(' | ')}`);
  });
}
