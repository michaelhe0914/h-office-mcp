import fs from 'fs';
import * as cheerio from 'cheerio';

const html = fs.readFileSync('deliver_date_page.html', 'utf-8');
const $ = cheerio.load(html);

const hsCustomers = [
  "嘉捷文化-高中台語",
  "旭如興業-高中台語",
  "效果升學-高中台語",
  "北區自營",
  "大亞文化-高中台語",
  "晟暘文教(湯曜瑋)",
  "展立事業-高中台語",
  "大中圖書-高中台語",
  "新興書局-高中台語",
  "總成資訊-高中台語",
  "文國書局-高中台語",
  "高雄天才出版社-高中台語",
  "學園社-高中台語",
  "建成書局",
  "宏家教育用品-高中台語",
  "大興書局-高中台語",
  "小城書局-高中台語",
  "斡耀通教學-高中台語",
  "文軒書局-高中台語",
  "大漢書局(板橋)-高中台語",
  "慈暉書局(台東)-高中台語",
  "南區自營",
  "大山書局-高中台語",
  "學明書局-高中台語",
  "文林書局-高中台語",
  "大雅書局-高中台語",
  "元品實業社-高中台語",
  "雨威實業社-高中台語",
  "超仁書局-高中台語",
  "中區自營",
  "新時代書局-高中台語",
  "清文實業",
  "神齊文化-高中台語",
  "偉文書局-高中台語",
  "吳振福",
  "達人書局-國中台.客",
  "均益書報社-高中台語",
  "方耀乾",
  "金安-台北分公司",
  "會計部"
];

console.log("=== Matching High School Customers to Deliver Date Links ===");
const links = {};
$('a').each((i, el) => {
  const text = $(el).text().trim();
  const href = $(el).attr('href');
  if (href.startsWith('/deliver_date/')) {
    links[text] = href;
  }
});

for (const hsCust of hsCustomers) {
  const baseName = hsCust.replace(/-高中台語$/, '').replace(/-國中台\.客$/, '');
  let matched = null;
  for (const [name, href] of Object.entries(links)) {
    if (name.includes(baseName) || baseName.includes(name)) {
      matched = { name, href };
      break;
    }
  }
  if (matched) {
    console.log(`HS: ${hsCust} => Matched Link: ${matched.name} (${matched.href})`);
  } else {
    console.log(`HS: ${hsCust} => ❌ No Match`);
  }
}
