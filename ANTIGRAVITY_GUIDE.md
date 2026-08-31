# 出書統計表 (h-office-mcp) Antigravity 使用與安裝指南

本指南說明如何在 **Antigravity AI** 中設定並使用「金安出版社出書統計表 MCP Server」。

---

## 🔐 前置準備：加入您的帳號與密碼

在使用 MCP 查詢前，**每位使用者必須設定自己的 h-office 系統帳號與密碼**。

1. **複製範本檔案**：
   在專案根目錄下，將 `.env.example` 複製為 `.env`：
   ```bash
   cp .env.example .env
   ```
   *（在 Windows 中亦可執行 `npm run setup` 自動建立）*

2. **編輯 `.env` 填入您的帳密**：
   使用文字編輯器打開 `.env` 檔案：
   ```ini
   H_OFFICE_EMAIL=your_email@gmail.com
   H_OFFICE_PASSWORD=your_password
   ```

> [!IMPORTANT]
> - `.env` 與登入產生之 Session `cookies.json` 皆已設定 Git 排除（`.gitignore`），個人密碼與權限**不會上傳**至 GitHub。
> - 當 Session 過期時，MCP 伺服器會使用您在 `.env` 中設定的帳密自動重新登入驗證。

---

## 🚀 步驟 1：安裝與編譯

1. 下載或 Clone 本專案：
   ```bash
   git clone https://github.com/michaelhe0914/h-office-mcp.git
   cd h-office-mcp
   ```

2. 安裝依賴與編譯：
   ```bash
   npm install
   npm run build
   ```

3. （選用）手動登入測試 / 雙重驗證 (2FA)：
   若您的 Google 帳號開啟了雙重驗證 (2FA)，建議先手動執行一次登入：
   ```bash
   npm run login
   ```
   這會啟動瀏覽器協助您完成 OAuth 登入並儲存 Session Cookie。

---

## ⚙️ 步驟 2：在 Antigravity 中設定 MCP 伺服器

您可以透過以下兩種方式之一，將本 MCP Server 整合至 Antigravity 中：

### 方法一：專案層級 / 設定檔設定 (推薦)
編輯或新增 Antigravity MCP 設定（如 `mcp.json` 或用戶端 MCP 設定）：

```json
{
  "mcpServers": {
    "h-office": {
      "command": "node",
      "args": [
        "C:/path/to/h-office-mcp/dist/index.js"
      ],
      "env": {
        "H_OFFICE_EMAIL": "your_email@gmail.com",
        "H_OFFICE_PASSWORD": "your_password"
      }
    }
  }
}
```
*(請將 `C:/path/to/h-office-mcp` 替換為您本機專案的實際絕對路徑)*

---

## 💬 步驟 3：在 Antigravity 中對話查詢

設定完成後，開啟 Antigravity 即可直接在對話框中以自然語言要求 AI 幫您查詢出書統計：

### 常用對話範例：

1. **檢查登入狀態**
   > 「幫我檢查 h-office 出書統計表的登入狀態。」

2. **列出可查詢的產品類別**
   > 「請列出出書統計表目前有哪些產品類別可以查詢。」

3. **查詢特定日期與產品類別**
   > 「幫我查詢 2026-5-25 到 2026-5-29 的 國中講義:複習講義(不含5-6) 銷售統計。」

4. **特定節日或新學期出貨統計**
   > 「請查詢 2025-9-26 到 2025-10-10 高中:複習講義 的出書統計表。」

---

## 🛠️ MCP 工具說明

| 工具名稱 | 說明 | 關鍵參數 |
| :--- | :--- | :--- |
| `check_login` | 檢查目前的 Session 是否有效及登入帳號。 | 無 |
| `list_product_classes` | 列出系統內所有可查詢的產品分類項目。 | 無 |
| `query_sales` | 查詢特定區間與類別之銷售統計表（格式化文字摘要）。 | `begin_date` (`YYYY-M-D`), `end_date` (`YYYY-M-D`), `product_class` (選填) |
| `query_sales_json` | 查詢銷售數據並回傳結構化 JSON，支援區域、業務員、經銷商與品名多重篩選及各維度加總。 | `begin_date`, `end_date`, `product_class`, `zone`, `sales_rep`, `customer`, `product` |
| `query_sales_raw` | 回傳原始 HTML 頁面（用於排錯除錯）。 | `begin_date`, `end_date`, `product_class` |

---

## ❓ 常見問題 FAQ

- **Q: 為什麼提示 `尚未登入`？**
  - **A**: 請確認 `.env` 檔案中已填入正確的 `H_OFFICE_EMAIL` 與 `H_OFFICE_PASSWORD`。若有 2FA，請執行 `npm run login` 完成驗證。

- **Q: 日期格式有什麼限制？**
  - **A**: 系統要求的日期格式為 `YYYY-M-D`（月份與日期不補零，例如 5 月 9 日寫作 `2026-5-9`）。

- **Q: 報表數據是否會洩漏？**
  - **A**: 所有查詢結果與 `.env`、`cookies.json` 皆已由 `.gitignore` 排除，絕不上傳雲端。
