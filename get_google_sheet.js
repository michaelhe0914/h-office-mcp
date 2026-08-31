import { request } from './dist/client.js';
import fs from 'fs';

async function run() {
  const url = "https://docs.google.com/a/king-an.com.tw/spreadsheet/ccc?key=0Ankow7Py-JZadGl4ZnJETlNQLUJ6SDdhNXlvR2VlaUE";
  console.log(`Fetching Google Sheet from ${url}...`);
  try {
    // Note: client.request is designed for h-office BASE_URL.
    // Let's use global fetch (available in Node 18+) or write a simple script to request it.
    const res = await fetch(url);
    const text = await res.text();
    fs.writeFileSync('google_sheet.html', text, 'utf-8');
    console.log("Saved google_sheet.html");
  } catch (e) {
    console.error("Error fetching Google Sheet:", e.message);
  }
}

run().catch(console.error);
