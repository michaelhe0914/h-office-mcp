# 金安出版社 h-office 系統查詢指南 (memory.md)

本文件記錄了金安出版社內部系統 `h-office` 出書統計表（銷售數據）的查詢方式、身分驗證機制、產品分類以及腳本執行說明，供日後開發或快速查詢時參考。

---

## 📌 系統基本資訊
*   **系統名稱**：金安出版社 h-office 辦公管理系統
*   **系統網址**：`https://h-office.king-an.com.tw:8082`
*   **憑證處理**：系統使用自簽憑證或較弱的 DH 密鑰，在 Node.js 中需特別設定：
    *   停用 TLS 未授權拒絕：`process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"`
    *   自訂 HTTPS Agent 啟用相容性：`ciphers: "DEFAULT@SECLEVEL=0"`, `minVersion: "TLSv1"`

---

## 🔐 身份驗證機制 (Google OAuth)
本系統登入流程基於 Google OAuth 2.0，驗證通過後會簽發 session cookie：`beaker.session.id`。

### 1. Cookies 儲存位置
*   本機檔案：`cookies.json`（位於專案根目錄）
*   檔案格式：
    ```json
    {
      "cookies": {
        "beaker.session.id": "your_session_id_here"
      },
      "updatedAt": "2026-05-30T03:13:18.314Z"
    }
    ```

### 2. 登入方法
當 Session 過期時，可使用以下兩種方式重新登入：

#### 方法一：自動化 Puppeteer 登入 (推薦)
1.  設定環境變數：
    *   Windows (CMD)：
        ```cmd
        set H_OFFICE_EMAIL=your_email@gmail.com
        set H_OFFICE_PASSWORD=your_password
        ```
    *   Windows (PowerShell)：
        ```powershell
        $env:H_OFFICE_EMAIL="your_email@gmail.com"
        $env:H_OFFICE_PASSWORD="your_password"
        ```
2.  執行登入腳本：
    ```bash
    npm run login
    ```
    *(此腳本會啟動 Chromium 瀏覽器，自動填入帳密並完成驗證後，將 Cookie 寫入 `cookies.json`)*

#### 方法二：手動複製 Cookie
1.  使用瀏覽器打開 `https://h-office.king-an.com.tw:8082/login` 並登入。
2.  按 F12 開啟開發者工具 $\rightarrow$ 切換至 **Application** 頁籤 $\rightarrow$ 展開左側 **Cookies**。
3.  找到 `beaker.session.id` 的值。
4.  直接編輯根目錄的 `cookies.json`，或設定環境變數 `H_OFFICE_COOKIE=beaker.session.id=xxx`。

---

## 📊 出書統計表查詢方式

### 1. 查詢端點與參數
*   **網址路徑**：`/sales`
*   **查詢方法**：`GET`
*   **常用參數**：
    | 參數名稱 | 格式 | 說明 | 範例 |
    | :--- | :--- | :--- | :--- |
    | `begin` | `YYYY-M-D` | 查詢區間的開始日期 | `2026-5-29` |
    | `end` | `YYYY-M-D` | 查詢區間的結束日期 | `2026-5-29` |
    | `product_class` | 字串 | 產品類別名稱（需完全相符） | `國中講義:複習講義(不含5-6)` |

> [!IMPORTANT]
> **日期格式限制**：系統所接受的日期格式為 `YYYY-M-D`（例如五月為 `5` 而非 `05`）。

---

## 📂 常見產品類別 (Product Classes)

查詢時，建議先使用 `list_product_classes` 工具或抓取頁面 Javascript 變數 `_product_classes` 取得最新類別。以下是常見與「複習」、「講義」相關的分類：

### 🎯 複習講義與書籍類
*   `國中講義:複習講義(不含5-6)` (最主要的國中複習講義分類)
*   `高中:複習講義` (高中複習講義分類)
*   `國中講義:新思維(不含5-6)` (複習講義性質)
*   `國中講義:輔導1-2` (單冊複習輔導講義)
*   `國中講義:考前30天`
*   `國中講義:3900題`

### 📝 複習考卷類
*   `國中考卷:複習卷-A卷`
*   `國中考卷:複習卷-B卷`
*   `國中考卷:複習卷-新思維`
*   `國中考卷:複習卷-半全冊`
*   `高中:複習考卷`

---

## 🛠️ MCP 伺服器工具說明
專案已打包為 MCP (Model Context Protocol) 伺服器，可在支援的 AI 用戶端（如 Cline）中直接調用以下工具：

1.  **`check_login`**：檢查目前的登入狀態與 beaker session 是否有效。
2.  **`list_product_classes`**：列出目前所有可供查詢的產品類別。
3.  **`query_sales`**：
    *   參數：`begin_date` (必填)、`end_date` (必填)、`product_class` (選填)
    *   功能：回傳格式化後的統計表（包含總計、分區統計與明細）。
4.  **`query_sales_raw`**：回傳原始 HTML 供偵錯分析頁面結構使用。

---

## 💻 快速查詢 Node.js 程式碼範例
若需要以程式碼直接查詢特定日期與種類，可參考以下快速寫法：

```javascript
import { request, loadCookies } from './dist/client.js';
import { parseSalesPage, formatSalesResults } from './dist/parser.js';

async function runQuery() {
  // 1. 載入本地存取的 beaker session cookie
  loadCookies();
  const BASE_URL = "https://h-office.king-an.com.tw:8082";

  // 2. 設定查詢條件
  const params = {
    begin: "2026-5-29",
    end: "2026-5-29",
    product_class: "國中講義:複習講義(不含5-6)"
  };

  // 3. 發送請求與解析資料
  console.log(`正在查詢 ${params.begin} [${params.product_class}] 統計資料...`);
  const res = await request(BASE_URL, "/sales", { params });
  const result = parseSalesPage(res.body);

  // 4. 印出格式化結果
  console.log(formatSalesResults(result));
}

runQuery().catch(console.error);
```

---

## 📋 查詢與回覆規範
為了確保查詢結果的精準性與簡潔度，在處理使用者查詢請求時，必須嚴格遵守以下準則：
1. **精準對齊範圍**：僅針對使用者明確指定的**日期**與**產品項目**（或類別）進行查詢與顯示。
2. **禁止主動擴大查詢**：除非使用者主動要求，否則**不要**主動提供非指定日期範圍（例如：整月、整週）或非指定項目之數據延伸，以避免資訊干擾，確保回覆精準無誤。

