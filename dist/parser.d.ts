/**
 * HTML parser for h-office sales statistics page.
 * Extracts structured data from server-rendered HTML tables using `td[name]` attributes.
 */
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
export declare function parseSalesPage(html: string): SalesResult;
/**
 * Extract product classes from the `_product_classes` JavaScript variable.
 */
export declare function parseProductClasses(html: string): ProductClass[];
/**
 * Check if the user is logged in from the HTML response.
 */
export declare function checkLoginStatus(html: string): {
    loggedIn: boolean;
    username?: string;
};
/**
 * Format sales results as a readable text summary.
 */
export declare function formatSalesResults(result: SalesResult): string;
