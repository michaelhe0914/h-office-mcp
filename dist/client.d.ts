/**
 * HTTP client for h-office system.
 * Handles self-signed SSL certificates and cookie-based authentication.
 */
/** Load cookies from disk */
export declare function loadCookies(): boolean;
/** Save cookies to disk */
export declare function saveCookies(cookies: Record<string, string>): void;
/** Get the cookie header string */
export declare function getCookieHeader(): string;
/** Check if we have a valid session */
export declare function hasSession(): boolean;
/** Get cookies file path */
export declare function getCookiesFilePath(): string;
/**
 * Make an HTTP(S) request to h-office, handling self-signed certs.
 */
export declare function request(baseUrl: string, urlPath: string, options?: {
    method?: string;
    params?: Record<string, string>;
    body?: string;
    followRedirects?: boolean;
    maxRedirects?: number;
}): Promise<{
    statusCode: number;
    headers: Record<string, any>;
    body: string;
}>;
