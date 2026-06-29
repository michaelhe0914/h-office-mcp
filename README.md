# 出書統計表工具 (h-office-mcp)

金安出版社內部 `h-office` 辦公系統「出書統計表」的查詢與報表工具。
核心是一個 **MCP (Model Context Protocol) Server**，可在支援的 AI 用戶端中查詢銷售統計；另附一組用於各區 / 各業務員出貨統計的查詢與 Excel 報表產生腳本。

> ⚠️ **內部使用**：本專案連線至公司內部系統並處理營業資料。請勿公開散布查詢結果、報表或登入憑證。

## 功能

- **MCP Server**（`src/` → `dist/`），提供工具：
  - `check_login` — 檢查 beaker session 是否有效
  - `list_product_classes` — 列出可查詢的產品類別
  - `query_sales` — 依日期區間與產品類別回傳格式化統計表
  - `query_sales_raw` — 回傳原始 HTML 供偵錯
- **報表腳本**：`query_*.js` 抓資料、`generate_*.js` 產生各區與各業務員的 Excel 統計表、`verify_*.js` 驗證結果。

## 技術

Node.js (ESM) · TypeScript · puppeteer（自動登入）· cheerio（HTML 解析）· exceljs / xlsx（報表）· zod

## 安裝

```bash
npm install
npm run build
```

## 登入

系統使用 Google OAuth，驗證後簽發 `beaker.session.id` cookie，存於專案根目錄的 `cookies.json`（**已 gitignore，不入庫**）。

```bash
# 自動化登入（Puppeteer）
export H_OFFICE_EMAIL="your_email@gmail.com"
export H_OFFICE_PASSWORD="your_password"
npm run login
```

或手動將瀏覽器中的 `beaker.session.id` 值填入 `cookies.json`。

## 使用

```bash
npm start          # 啟動 MCP server
node query_*.js    # 執行特定查詢/報表腳本
```

詳細系統說明見 [`memory.md`](memory.md) 與 [`國中講義單冊出貨統計_運作方式.md`](國中講義單冊出貨統計_運作方式.md)。

## 注意事項

- `cookies.json` 與所有查詢結果（`query_results_*.json`）、Excel 報表（`Output*/`、`*.xlsx`）皆已排除於版控之外，內含營業機密與個資。
- 帳號密碼僅透過環境變數提供，不寫入程式碼。
