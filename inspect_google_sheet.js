import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('google_sheet.html', 'utf-8');
const $ = cheerio.load(html);

console.log("=== Title of Google Sheet ===");
console.log($('title').text());

console.log("\n=== Content Length ===");
console.log(html.length);

console.log("\n=== Checking for login redirects ===");
if (html.includes('service=wise') || html.includes('accounts.google.com') || html.includes('Sign in')) {
  console.log("Redirected to Google Accounts login page.");
} else {
  console.log("Looks like we got the sheet or something else.");
  console.log(html.substring(0, 1000));
}
