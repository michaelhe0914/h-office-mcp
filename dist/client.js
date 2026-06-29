/**
 * HTTP client for h-office system.
 * Handles self-signed SSL certificates and cookie-based authentication.
 */
import https from "node:https";
import { URL, fileURLToPath } from "node:url";
import fs from "node:fs";
import path from "node:path";
/**
 * Custom HTTPS agent that tolerates the h-office server's weak DH parameters
 * and self-signed certificate.
 */
const agent = new https.Agent({
    rejectUnauthorized: false,
    // Allow weak DH keys used by the server
    ciphers: "DEFAULT@SECLEVEL=0",
    minVersion: "TLSv1",
});
const COOKIES_FILE = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "cookies.json");
let cookieStore = { cookies: {}, updatedAt: "" };
/** Load cookies from disk */
export function loadCookies() {
    try {
        if (fs.existsSync(COOKIES_FILE)) {
            const data = JSON.parse(fs.readFileSync(COOKIES_FILE, "utf-8"));
            cookieStore = data;
            return Object.keys(cookieStore.cookies).length > 0;
        }
    }
    catch {
        // ignore
    }
    return false;
}
/** Save cookies to disk */
export function saveCookies(cookies) {
    cookieStore = {
        cookies: { ...cookieStore.cookies, ...cookies },
        updatedAt: new Date().toISOString(),
    };
    fs.writeFileSync(COOKIES_FILE, JSON.stringify(cookieStore, null, 2), "utf-8");
}
/** Get the cookie header string */
export function getCookieHeader() {
    return Object.entries(cookieStore.cookies)
        .map(([k, v]) => `${k}=${v}`)
        .join("; ");
}
/** Check if we have a valid session */
export function hasSession() {
    return "beaker.session.id" in cookieStore.cookies;
}
/** Get cookies file path */
export function getCookiesFilePath() {
    return COOKIES_FILE;
}
/** Parse Set-Cookie headers */
function parseSetCookies(headers) {
    const result = {};
    if (!headers)
        return result;
    const arr = Array.isArray(headers) ? headers : [headers];
    for (const h of arr) {
        const match = h.match(/^([^=]+)=([^;]*)/);
        if (match) {
            result[match[1].trim()] = match[2].trim();
        }
    }
    return result;
}
/**
 * Make an HTTP(S) request to h-office, handling self-signed certs.
 */
export function request(baseUrl, urlPath, options = {}) {
    const { method = "GET", params, body, followRedirects = false, maxRedirects = 5 } = options;
    return new Promise((resolve, reject) => {
        let fullUrl = `${baseUrl}${urlPath}`;
        if (params && method === "GET") {
            const qs = new URLSearchParams(params).toString();
            if (qs)
                fullUrl += `?${qs}`;
        }
        const parsed = new URL(fullUrl);
        const reqOptions = {
            hostname: parsed.hostname,
            port: parsed.port || 443,
            path: parsed.pathname + parsed.search,
            method,
            agent,
            headers: {
                Cookie: getCookieHeader(),
                "User-Agent": "h-office-mcp/1.0",
            },
        };
        if (body && method === "POST") {
            reqOptions.headers["Content-Type"] = "application/x-www-form-urlencoded";
            reqOptions.headers["Content-Length"] = Buffer.byteLength(body);
        }
        const req = https.request(reqOptions, (res) => {
            // Parse and save any cookies from response
            const setCookies = parseSetCookies(res.headers["set-cookie"]);
            if (Object.keys(setCookies).length > 0) {
                saveCookies(setCookies);
            }
            // Handle redirects
            if (followRedirects &&
                maxRedirects > 0 &&
                res.statusCode &&
                [301, 302, 303, 307, 308].includes(res.statusCode) &&
                res.headers.location) {
                const location = res.headers.location;
                const redirectUrl = location.startsWith("http")
                    ? new URL(location)
                    : new URL(location, fullUrl);
                // Only follow redirects to the same host
                if (redirectUrl.hostname === parsed.hostname) {
                    request(baseUrl, redirectUrl.pathname + redirectUrl.search, {
                        method: res.statusCode === 303 ? "GET" : method,
                        followRedirects: true,
                        maxRedirects: maxRedirects - 1,
                    })
                        .then(resolve)
                        .catch(reject);
                    return;
                }
            }
            const chunks = [];
            res.on("data", (chunk) => {
                chunks.push(chunk);
            });
            res.on("end", () => {
                const body = Buffer.concat(chunks).toString("utf-8");
                resolve({
                    statusCode: res.statusCode || 0,
                    headers: res.headers,
                    body,
                });
            });
        });
        req.on("error", reject);
        if (body && method === "POST") {
            req.write(body);
        }
        req.end();
    });
}
