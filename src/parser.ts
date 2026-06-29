/**
 * HTML parser for h-office sales statistics page.
 * Extracts structured data from server-rendered HTML tables using `td[name]` attributes.
 */

import * as cheerio from "cheerio";

export interface SalesRecord {
  /** 區域 */
  zone: string;
  /** 業務人員 */
  sales: string;
  /** 經銷商 */
  customer: string;
  /** 產品名稱 */
  product: string;
  /** 訂量 */
  qty: number;
  /** 退量 */
  rtn_qty: number;
}

export interface ProductClass {
  /** 類別全名 (e.g., "國中講義:複習講義(不含5-6)") */
  name: string;
  /** 標籤 */
  tags: string[];
}

export interface SalesResult {
  /** 是否已登入 */
  isLoggedIn: boolean;
  /** 登入使用者 */
  username?: string;
  /** 查詢期間 */
  period?: string;
  /** 查詢的產品類別 */
  productClass?: string;
  /** 資料列 */
  records: SalesRecord[];
  /** 可用的產品類別 */
  productClasses: ProductClass[];
  /** 總訂量 */
  totalQty: number;
  /** 總退量 */
  totalRtnQty: number;
}

/**
 * Parse the sales page HTML and extract structured data.
 * The table uses `td[name]` attributes: zone, sales, customer, product, qty, rtn_qty
 */
export function parseSalesPage(html: string): SalesResult {
  const $ = cheerio.load(html);

  // Check login status
  const userSpan = $("span.user").text().trim();
  const isLoggedIn = !userSpan.includes("尚未登入");
  const username = isLoggedIn
    ? userSpan.replace(/\s*登出.*$/, "").trim()
    : undefined;

  // Extract product classes from JavaScript variable
  const productClasses = parseProductClasses(html);

  // Extract current product class from hidden input
  const currentProductClass =
    $("input[name='product_class']").val()?.toString() || undefined;

  // Parse data rows from the table
  const records: SalesRecord[] = [];
  $("table.list tbody tr.data").each((_, row) => {
    const $row = $(row);
    const zone = $row.find("td[name='zone']").text().trim();
    const sales = $row.find("td[name='sales']").text().trim();
    const customer = $row.find("td[name='customer']").text().trim();
    const product = $row.find("td[name='product']").text().trim();
    const qty = parseInt($row.find("td[name='qty']").text().trim(), 10) || 0;
    const rtn_qty =
      parseInt($row.find("td[name='rtn_qty']").text().trim(), 10) || 0;

    if (product) {
      records.push({ zone, sales, customer, product, qty, rtn_qty });
    }
  });

  // Calculate totals
  const totalQty = records.reduce((sum, r) => sum + r.qty, 0);
  const totalRtnQty = records.reduce((sum, r) => sum + r.rtn_qty, 0);

  return {
    isLoggedIn,
    username,
    productClass: currentProductClass,
    records,
    productClasses,
    totalQty,
    totalRtnQty,
  };
}

/**
 * Extract product classes from the `_product_classes` JavaScript variable.
 */
export function parseProductClasses(html: string): ProductClass[] {
  const match = html.match(/var _product_classes\s*=\s*(\[[\s\S]*?\]);/);
  if (!match) return [];

  try {
    const parsed = JSON.parse(match[1]);
    return parsed.map((item: any) => ({
      name: item.name || "",
      tags: item.tags || [],
    }));
  } catch {
    return [];
  }
}

/**
 * Check if the user is logged in from the HTML response.
 */
export function checkLoginStatus(html: string): {
  loggedIn: boolean;
  username?: string;
} {
  const $ = cheerio.load(html);
  const userSpan = $("span.user").text().trim();

  if (userSpan.includes("尚未登入")) {
    return { loggedIn: false };
  }

  const username = userSpan.replace(/\s*登出.*$/, "").trim() || undefined;
  return { loggedIn: true, username };
}

/**
 * Format sales results as a readable text summary.
 */
export function formatSalesResults(result: SalesResult): string {
  const lines: string[] = [];

  lines.push(`📊 出書統計表`);
  if (result.username) lines.push(`👤 使用者：${result.username}`);
  if (result.period) lines.push(`📅 期間：${result.period}`);
  if (result.productClass) lines.push(`📁 類別：${result.productClass}`);
  lines.push("");

  if (!result.isLoggedIn) {
    lines.push("⚠️ 尚未登入，無法取得資料。請先執行登入程序。");
    return lines.join("\n");
  }

  if (result.records.length === 0) {
    if (!result.productClass) {
      lines.push("📋 請指定產品類別 (product_class) 來查詢。");
      lines.push("   可使用 list_product_classes 工具查看可用的類別。");
    } else {
      lines.push("（查無資料）");
    }
    return lines.join("\n");
  }

  // Summary stats
  lines.push(
    `📈 總計：訂量 ${result.totalQty.toLocaleString()} | 退量 ${result.totalRtnQty.toLocaleString()} | 退書率 ${result.totalQty > 0 ? ((result.totalRtnQty / (result.totalQty + result.totalRtnQty)) * 100).toFixed(1) : 0}%`
  );
  lines.push(`📝 共 ${result.records.length} 筆資料`);
  lines.push("");

  // Group by zone for summary
  const byZone = new Map<string, { qty: number; rtn: number; count: number }>();
  for (const r of result.records) {
    const z = byZone.get(r.zone) || { qty: 0, rtn: 0, count: 0 };
    z.qty += r.qty;
    z.rtn += r.rtn_qty;
    z.count++;
    byZone.set(r.zone, z);
  }

  lines.push("=== 各區域統計 ===");
  for (const [zone, data] of byZone) {
    lines.push(
      `  ${zone}：訂量 ${data.qty.toLocaleString()} | 退量 ${data.rtn.toLocaleString()} | ${data.count} 筆`
    );
  }
  lines.push("");

  // Detail table (limit to avoid overly long output)
  const MAX_ROWS = 200;
  lines.push("=== 明細 ===");
  lines.push("區域 | 業務 | 經銷 | 產品 | 訂量 | 退量");
  lines.push("-----|------|------|------|------|------");

  const displayRecords = result.records.slice(0, MAX_ROWS);
  for (const r of displayRecords) {
    lines.push(
      `${r.zone} | ${r.sales} | ${r.customer} | ${r.product} | ${r.qty} | ${r.rtn_qty}`
    );
  }

  if (result.records.length > MAX_ROWS) {
    lines.push(
      `... 還有 ${result.records.length - MAX_ROWS} 筆資料未顯示`
    );
  }

  return lines.join("\n");
}
