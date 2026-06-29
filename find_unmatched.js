import fs from 'fs';

const mainRecords = JSON.parse(fs.readFileSync('query_results_north_115.json', 'utf-8'));

const unmatched = mainRecords.filter(r => {
  if (r.productClass === "國中講義:新思維5-6") return false;
  if (r.productClass === "國中講義:新思維(不含5-6)") return false;
  
  const isSx = r.product.includes("雙向");
  const is735 = r.product.includes("735");
  const isTheme = r.product.includes("主題讚");
  return !(isSx || is735 || isTheme);
});

console.log("Unmatched records:", unmatched.length);
const uniqueProducts = [...new Set(unmatched.map(r => r.product))];
console.log("Unique unmatched products:", uniqueProducts);
