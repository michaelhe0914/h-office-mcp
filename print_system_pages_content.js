import * as fs from 'fs';
import * as cheerio from 'cheerio';

for (const filename of ['./check_raw.html', './setting_raw.html']) {
  if (fs.existsSync(filename)) {
    const html = fs.readFileSync(filename, 'utf8');
    const $ = cheerio.load(html);
    console.log(`\n==================================================`);
    console.log(`Content of ${filename}:`);
    console.log(`==================================================`);
    
    // Print the main div or body text
    const text = $('#content').text().trim() || $('body').text().trim();
    console.log(text.substring(0, 1500));
    
    // Print any links
    const links = [];
    $('a').each((_, elem) => {
      const t = $(elem).text().trim();
      const href = $(elem).attr('href');
      if (href && !href.startsWith('/') && href !== '#') {
        links.push(`${t} (${href})`);
      } else if (href) {
        links.push(`${t} (${href})`);
      }
    });
    if (links.length > 0) {
      console.log("\nLinks found:");
      console.log(links.slice(0, 20));
    }
  }
}
