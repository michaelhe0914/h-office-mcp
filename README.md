# 出書統計表工具 (h-office-mcp)

金安出版社內部 `h-office` 辦公系統「出書統計表」的查詢與報表工具。
核心是一個 **MCP (Model Context Protocol) Server**，可在支援的 AI 用戶端（如 **Google Antigravity**）中以自然語言查詢銷售統計；另附一組用於各區 / 各業務員出貨統計的查詢與 Excel 報表產生腳本。

> ⚠️ **內部使用**：本專案連線至公司內部系統並處理營業資料。請勿公開散布查詢結果、報表或登入憑證。

---

## 📖 Antigravity 快速入門指南

要在 **Google Antigravity** 中讓其他成員也能輸入自己的帳號密碼對話查詢，請參閱詳細教學：
👉 [**ANTIGRAVITY_GUIDE.md**](ANTIGRAVITY_GUIDE.md)

---

## 🛠️ 功能

- **MCP Server**（`src/` → `dist/`），提供工具：
  - `check_login` — 檢查 beaker session 是否有效與登入身分
  - `list_product_classes` — 列出可查詢的產品類別
  - `query_sales` — 依日期區間與產品類別回傳格式化統計表
  - `query_sales_json` — 依日期區間、區域、業務員等回傳結構化 JSON 與統計數據
  - `query_sales_raw` — 回傳原始 HTML 供偵錯
- **報表腳本**：`query_*.js` 抓資料、`generate_*.js` 產生各區與各業務員的 Excel 統計表、`verify_*.js` 驗證結果。

---

## 🚀 快速安裝與設定

1. **複製範本並設定個人的帳號密碼**：
   ```bash
   cp .env.example .env
   ```
   *（在 WindowsCMD/PowerShell 亦可執行 `npm run setup`）*

   編輯 `.env` 檔案：
   ```ini
   H_OFFICE_EMAIL=your_email@gmail.com
   H_OFFICE_PASSWORD=your_password
   ```

2. **安裝套件與編譯**：
   ```bash
   npm install
   npm run build
   ```

3. **執行登入與驗證**：
   ```bash
   npm run login
   ```

---

## 🔐 登入機制與隱私保護

- 系統登入使用 Google OAuth，驗證通過後會簽發 `beaker.session.id` Cookie 存於 `cookies.json`。
- **所有機密檔案（`.env`、`cookies.json`、`Output*/`、`*.xlsx`、`query_results_*.json`）皆已包含於 `.gitignore` 中，保證個人密碼與營業機密絕不上傳版控。**

---

## 💻 使用方式

```bash
npm start          # 啟動 MCP server (Stdio)
npm run login      # 執行 Puppeteer 自動/手動登入
node query_*.js    # 執行特定查詢/報表腳本
```

詳細系統規格與運作細節說明見 [`memory.md`](memory.md)、[`ANTIGRAVITY_GUIDE.md`](ANTIGRAVITY_GUIDE.md) 與 [`國中講義單冊出貨統計_運作方式.md`](國中講義單冊出貨統計_運作方式.md)。
