// List all product classes from h-office
import { request, loadCookies } from './dist/client.js';
import { parseSalesPage } from './dist/parser.js';

async function run() {
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";
  const res = await request(BASE_URL, "/sales");
  const result = parseSalesPage(res.body);
  console.log("=== Product Classes ===");
  result.productClasses.forEach(pc => {
    console.log(`- ${pc.name} (Tags: ${pc.tags.join(', ')})`);
  });
}

run().catch(console.error);
