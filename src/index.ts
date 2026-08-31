#!/usr/bin/env node

/**
 * MCP Server for King-An Publishing (金安出版社) h-office system.
 *
 * Provides tools to query sales statistics (出書統計表) from the internal
 * office management system at https://h-office.king-an.com.tw:8082
 *
 * Authentication: Google OAuth 2.0 → beaker.session.id cookie
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  request,
  loadCookies,
  hasSession,
  saveCookies,
  getCookiesFilePath,
  loadEnv,
} from "./client.js";
import {
  parseSalesPage,
  parseProductClasses,
  formatSalesResults,
  checkLoginStatus,
} from "./parser.js";
import { loginWithPuppeteer } from "./login.js";

// Load .env configuration on startup
loadEnv();

// ─── Configuration ─────────────────────────────────────────────────
const BASE_URL =
  process.env.H_OFFICE_BASE_URL || "https://h-office.king-an.com.tw:8082";

// Suppress Node.js TLS self-signed cert warnings
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

// ─── Server Setup ──────────────────────────────────────────────────
const server = new McpServer({
  name: "h-office",
  version: "1.0.0",
});

// ─── Helper: ensure login ──────────────────────────────────────────
async function ensureLoggedIn(): Promise<{ ok: boolean; message: string }> {
  loadEnv();

  if (!hasSession()) {
    // Try loading cookies from file
    loadCookies();
    if (!hasSession()) {
      // Check if cookie is provided via env var
      const envCookie = process.env.H_OFFICE_COOKIE;
      if (envCookie) {
        const cookies: Record<string, string> = {};
        envCookie.split(";").forEach((part) => {
          const [k, ...v] = part.trim().split("=");
          if (k) cookies[k.trim()] = v.join("=").trim();
        });
        saveCookies(cookies);
      }
    }
  }

  // Attempt auto-login if credentials are configured but no valid session
  if (!hasSession() && process.env.H_OFFICE_EMAIL && process.env.H_OFFICE_PASSWORD) {
    console.error("🔄 Attempting automatic login with configured environment credentials...");
    try {
      const success = await loginWithPuppeteer();
      if (success) {
        loadCookies();
      }
    } catch (e) {
      console.error("Auto-login error:", e);
    }
  }

  if (!hasSession()) {
    return {
      ok: false,
      message:
        `尚未登入。請先在專案根目錄的 .env 檔案中設定您的帳號密碼：\n\n` +
        `  H_OFFICE_EMAIL=your_email@gmail.com\n` +
        `  H_OFFICE_PASSWORD=your_password\n\n` +
        `或直接執行登入指令：\n` +
        `  npm run login\n\n` +
        `Cookies 檔案路徑：${getCookiesFilePath()}`,
    };
  }

  // Verify the session is still valid
  try {
    const res = await request(BASE_URL, "/");
    const status = checkLoginStatus(res.body);
    if (!status.loggedIn) {
      // If session expired and credentials present, retry auto-login
      if (process.env.H_OFFICE_EMAIL && process.env.H_OFFICE_PASSWORD) {
        console.error("🔄 Session expired. Re-attempting Puppeteer login...");
        const relogged = await loginWithPuppeteer();
        if (relogged) {
          loadCookies();
          const retryRes = await request(BASE_URL, "/");
          const retryStatus = checkLoginStatus(retryRes.body);
          if (retryStatus.loggedIn) {
            return { ok: true, message: `已登入：${retryStatus.username || "OK"}` };
          }
        }
      }
      return {
        ok: false,
        message:
          "Session 已過期。請重新設定 `.env` 密碼或執行 `npm run login`。\n" +
          `Cookies 檔案路徑：${getCookiesFilePath()}`,
      };
    }
    return { ok: true, message: `已登入：${status.username || "OK"}` };
  } catch (err) {
    return {
      ok: false,
      message: `連線失敗：${err instanceof Error ? err.message : String(err)}`,
    };
  }
}


// ─── Tool: check_login ─────────────────────────────────────────────
server.tool(
  "check_login",
  "檢查 h-office 系統的登入狀態。",
  {},
  async () => {
    const result = await ensureLoggedIn();
    return {
      content: [
        {
          type: "text" as const,
          text: result.ok
            ? `✅ ${result.message}`
            : `❌ ${result.message}`,
        },
      ],
    };
  }
);

// ─── Tool: list_product_classes ────────────────────────────────────
server.tool(
  "list_product_classes",
  "列出出書統計表可用的產品類別。需要先登入。",
  {},
  async () => {
    const auth = await ensureLoggedIn();
    if (!auth.ok) {
      return {
        content: [{ type: "text" as const, text: auth.message }],
        isError: true,
      };
    }

    try {
      const res = await request(BASE_URL, "/sales");
      const classes = parseProductClasses(res.body);

      if (classes.length === 0) {
        return {
          content: [
            {
              type: "text" as const,
              text: "未找到產品類別。可能需要重新登入或頁面結構已變更。",
            },
          ],
        };
      }

      return {
        content: [
          {
            type: "text" as const,
            text:
              `📁 可用的產品類別 (共 ${classes.length} 個)：\n\n` +
              classes.map((c, i) => `${i + 1}. ${c}`).join("\n"),
          },
        ],
      };
    } catch (err) {
      return {
        content: [
          {
            type: "text" as const,
            text: `查詢失敗：${err instanceof Error ? err.message : String(err)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ─── Tool: query_sales ─────────────────────────────────────────────
server.tool(
  "query_sales",
  "查詢金安出版社的出書統計表。支援按日期區間和產品類別篩選。日期格式為 YYYY-M-D（例如 2025-3-25）。",
  {
    begin_date: z
      .string()
      .describe("開始日期，格式 YYYY-M-D，例如 2025-3-25（上學期起始）或 2024-9-26（下學期起始）"),
    end_date: z
      .string()
      .describe("結束日期，格式 YYYY-M-D，例如 2025-5-30"),
    product_class: z
      .string()
      .optional()
      .describe("產品類別（可選）。使用 list_product_classes 工具查看可用類別。"),
  },
  async ({ begin_date, end_date, product_class }) => {
    const auth = await ensureLoggedIn();
    if (!auth.ok) {
      return {
        content: [{ type: "text" as const, text: auth.message }],
        isError: true,
      };
    }

    try {
      const params: Record<string, string> = {
        begin: begin_date,
        end: end_date,
      };
      if (product_class) {
        params.product_class = product_class;
      }

      const res = await request(BASE_URL, "/sales", { params });

      const result = parseSalesPage(res.body);
      result.period = `${begin_date} ~ ${end_date}`;
      result.productClass = product_class;

      const formatted = formatSalesResults(result);

      return {
        content: [{ type: "text" as const, text: formatted }],
      };
    } catch (err) {
      return {
        content: [
          {
            type: "text" as const,
            text: `查詢失敗：${err instanceof Error ? err.message : String(err)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ─── Tool: query_sales_json ────────────────────────────────────────
server.tool(
  "query_sales_json",
  "查詢金安出版社出書統計表並回傳結構化 JSON 資料。支援日期區間、產品類別、區域、業務員、經銷商與產品名稱篩選，並附帶各維度統計彙總。",
  {
    begin_date: z
      .string()
      .describe("開始日期，格式 YYYY-M-D，例如 2025-3-25"),
    end_date: z
      .string()
      .describe("結束日期，格式 YYYY-M-D，例如 2025-5-30"),
    product_class: z
      .string()
      .optional()
      .describe("產品類別（可選）。可透過 list_product_classes 查看。"),
    zone: z
      .string()
      .optional()
      .describe("篩選區域（可選，如：北區、中區、南區）"),
    sales_rep: z
      .string()
      .optional()
      .describe("篩選業務人員（可選，如：何光傑、李敏豪、蔡榮訓等）"),
    customer: z
      .string()
      .optional()
      .describe("篩選經銷商 / 書局名稱（可選）"),
    product: z
      .string()
      .optional()
      .describe("篩選產品名稱關鍵字（可選）"),
  },
  async ({ begin_date, end_date, product_class, zone, sales_rep, customer, product }) => {
    const auth = await ensureLoggedIn();
    if (!auth.ok) {
      return {
        content: [{ type: "text" as const, text: auth.message }],
        isError: true,
      };
    }

    try {
      const params: Record<string, string> = {
        begin: begin_date,
        end: end_date,
      };
      if (product_class) {
        params.product_class = product_class;
      }

      const res = await request(BASE_URL, "/sales", { params });
      const result = parseSalesPage(res.body);

      let filtered = result.records;
      if (zone) {
        filtered = filtered.filter((r) => r.zone.includes(zone));
      }
      if (sales_rep) {
        filtered = filtered.filter((r) => r.sales.includes(sales_rep));
      }
      if (customer) {
        filtered = filtered.filter((r) => r.customer.includes(customer));
      }
      if (product) {
        filtered = filtered.filter((r) => r.product.includes(product));
      }

      const totalQty = filtered.reduce((sum, r) => sum + r.qty, 0);
      const totalRtnQty = filtered.reduce((sum, r) => sum + r.rtn_qty, 0);
      const netQty = totalQty - totalRtnQty;

      // Aggregations
      const byZone: Record<string, { qty: number; rtn_qty: number; net_qty: number }> = {};
      const bySalesRep: Record<string, { qty: number; rtn_qty: number; net_qty: number }> = {};
      const byProduct: Record<string, { qty: number; rtn_qty: number; net_qty: number }> = {};

      for (const r of filtered) {
        // zone
        if (!byZone[r.zone]) byZone[r.zone] = { qty: 0, rtn_qty: 0, net_qty: 0 };
        byZone[r.zone].qty += r.qty;
        byZone[r.zone].rtn_qty += r.rtn_qty;
        byZone[r.zone].net_qty += r.qty - r.rtn_qty;

        // sales
        if (!bySalesRep[r.sales]) bySalesRep[r.sales] = { qty: 0, rtn_qty: 0, net_qty: 0 };
        bySalesRep[r.sales].qty += r.qty;
        bySalesRep[r.sales].rtn_qty += r.rtn_qty;
        bySalesRep[r.sales].net_qty += r.qty - r.rtn_qty;

        // product
        if (!byProduct[r.product]) byProduct[r.product] = { qty: 0, rtn_qty: 0, net_qty: 0 };
        byProduct[r.product].qty += r.qty;
        byProduct[r.product].rtn_qty += r.rtn_qty;
        byProduct[r.product].net_qty += r.qty - r.rtn_qty;
      }

      const outputData = {
        period: `${begin_date} ~ ${end_date}`,
        productClass: product_class || "全部",
        filters: { zone, sales_rep, customer, product },
        summary: {
          totalQty,
          totalRtnQty,
          netQty,
          recordCount: filtered.length,
        },
        byZone,
        bySalesRep,
        byProduct,
        records: filtered,
      };

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(outputData, null, 2),
          },
        ],
      };
    } catch (err) {
      return {
        content: [
          {
            type: "text" as const,
            text: `查詢失敗：${err instanceof Error ? err.message : String(err)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ─── Tool: query_sales_raw ─────────────────────────────────────────
server.tool(
  "query_sales_raw",
  "查詢出書統計表並回傳原始 HTML（用於偵錯或分析頁面結構）。",
  {
    begin_date: z.string().describe("開始日期，格式 YYYY-M-D"),
    end_date: z.string().describe("結束日期，格式 YYYY-M-D"),
    product_class: z.string().optional().describe("產品類別（可選）"),
  },
  async ({ begin_date, end_date, product_class }) => {
    const auth = await ensureLoggedIn();
    if (!auth.ok) {
      return {
        content: [{ type: "text" as const, text: auth.message }],
        isError: true,
      };
    }

    try {
      const params: Record<string, string> = {
        begin: begin_date,
        end: end_date,
      };
      if (product_class) {
        params.product_class = product_class;
      }

      const res = await request(BASE_URL, "/sales", { params });

      // Extract just the content div
      const contentMatch = res.body.match(
        /<div id='content'>([\s\S]*?)<\/div>\s*<\/div>\s*<\/body>/
      );
      const content = contentMatch ? contentMatch[1] : res.body;

      return {
        content: [
          {
            type: "text" as const,
            text: `HTTP ${res.statusCode}\n\n${content.substring(0, 8000)}`,
          },
        ],
      };
    } catch (err) {
      return {
        content: [
          {
            type: "text" as const,
            text: `查詢失敗：${err instanceof Error ? err.message : String(err)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ─── Start Server ──────────────────────────────────────────────────
async function main() {
  // Try to load saved cookies on startup
  loadCookies();

  const transport = new StdioServerTransport();
  await server.connect(transport);

  // Log to stderr (stdout is used for MCP protocol)
  console.error("🚀 h-office MCP Server started");
  console.error(`   Base URL: ${BASE_URL}`);
  console.error(`   Session: ${hasSession() ? "✅ found" : "❌ not found"}`);
  console.error(`   Cookies file: ${getCookiesFilePath()}`);
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
