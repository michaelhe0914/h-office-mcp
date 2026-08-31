import { request, loadCookies } from './dist/client.js';
import fs from 'fs';
import * as cheerio from 'cheerio';

const dealers = {
  "042032": "墨莉花企業社",
  "043014": "碩成企業社",
  "042031": "鴻良有限公司",
  "042002": "弘茂圖書社",
  "049901": "北區自營",
  "043013": "和平社",
  "040005": "環華文教",
  "040051": "新時代書局",
  "041033": "捷客有限公司",
  "041032": "合翔圖書",
  "042033": "誠信文教",
  "041023": "達人書局",
  "042028": "清文實業",
  "043032": "大興書局",
  "043037": "六育有限公司",
  "041031": "廷徽企業社",
  "040027": "天立社",
  "049054": "光文社",
  "041025": "超仁書局",
  "043038": "平大教育"
};

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  
  console.log("=== Querying deliver dates for all North Region dealers ===");
  for (const [id, name] of Object.entries(dealers)) {
    const url = `/deliver_date/${id}`;
    try {
      const res = await request(BASE_URL, url);
      const $ = cheerio.load(res.body);
      
      let rowsCount = 0;
      let hsCount = 0;
      
      $('table tr').slice(1).each((i, row) => {
        const cells = $(row).find('td').map((j, cell) => $(cell).text().trim()).get();
        if (cells.length > 0) {
          rowsCount++;
          const schoolName = cells[2] || "";
          if (
            schoolName.includes("高中") || 
            schoolName.includes("中學") || 
            schoolName.includes("高職") || 
            schoolName.includes("家商") || 
            schoolName.includes("商工") || 
            schoolName.includes("工農") || 
            schoolName.includes("中道")
          ) {
            console.log(`  [${name} (${id})] School: ${schoolName} | Status: ${cells[3]}`);
            hsCount++;
          }
        }
      });
      
      if (hsCount > 0) {
        console.log(`  => ${name} (${id}) has ${hsCount} high schools (out of ${rowsCount} schools).`);
      }
    } catch (e) {
      console.error(`  Error querying ${name} (${id}):`, e.message);
    }
  }
}

run().catch(console.error);
