/**
 * HTTP client for h-office system.
 * Handles self-signed SSL certificates and cookie-based authentication.
 */

import https from "node:https";
import tls from "node:tls";
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
  minVersion: "TLSv1" as tls.SecureVersion,
});

const COOKIES_FILE = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "cookies.json"
);

const ENV_FILE = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  ".env"
);

/** Load .env file into process.env */
export function loadEnv(): void {
  try {
    if (fs.existsSync(ENV_FILE)) {
      const content = fs.readFileSync(ENV_FILE, "utf-8");
      const lines = content.split("\n");
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith("#")) continue;
        const eqIdx = trimmed.indexOf("=");
        if (eqIdx > 0) {
          const key = trimmed.slice(0, eqIdx).trim();
          let val = trimmed.slice(eqIdx + 1).trim();
          if (
            (val.startsWith('"') && val.endsWith('"')) ||
            (val.startsWith("'") && val.endsWith("'"))
          ) {
            val = val.slice(1, -1);
          }
          if (!process.env[key]) {
            process.env[key] = val;
          }
        }
      }
    }
  } catch {
    // ignore
  }
}


interface CookieStore {
  cookies: Record<string, string>;
  updatedAt: string;
}

let cookieStore: CookieStore = { cookies: {}, updatedAt: "" };

/** Load cookies from disk */
export function loadCookies(): boolean {
  try {
    if (fs.existsSync(COOKIES_FILE)) {
      const data = JSON.parse(fs.readFileSync(COOKIES_FILE, "utf-8"));
      cookieStore = data;
      return Object.keys(cookieStore.cookies).length > 0;
    }
  } catch {
    // ignore
  }
  return false;
}

/** Save cookies to disk */
export function saveCookies(cookies: Record<string, string>): void {
  cookieStore = {
    cookies: { ...cookieStore.cookies, ...cookies },
    updatedAt: new Date().toISOString(),
  };
  fs.writeFileSync(COOKIES_FILE, JSON.stringify(cookieStore, null, 2), "utf-8");
}

/** Get the cookie header string */
export function getCookieHeader(): string {
  return Object.entries(cookieStore.cookies)
    .map(([k, v]) => `${k}=${v}`)
    .join("; ");
}

/** Check if we have a valid session */
export function hasSession(): boolean {
  return "beaker.session.id" in cookieStore.cookies;
}

/** Get cookies file path */
export function getCookiesFilePath(): string {
  return COOKIES_FILE;
}

/** Parse Set-Cookie headers */
function parseSetCookies(headers: string | string[] | undefined): Record<string, string> {
  const result: Record<string, string> = {};
  if (!headers) return result;

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
export function request(
  baseUrl: string,
  urlPath: string,
  options: {
    method?: string;
    params?: Record<string, string>;
    body?: string;
    followRedirects?: boolean;
    maxRedirects?: number;
  } = {}
): Promise<{ statusCode: number; headers: Record<string, any>; body: string }> {
  const { method = "GET", params, body, followRedirects = false, maxRedirects = 5 } = options;

  return new Promise((resolve, reject) => {
    let fullUrl = `${baseUrl}${urlPath}`;
    if (params && method === "GET") {
      const qs = new URLSearchParams(params).toString();
      if (qs) fullUrl += `?${qs}`;
    }

    const parsed = new URL(fullUrl);

    const reqOptions: https.RequestOptions = {
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
      (reqOptions.headers as any)["Content-Type"] = "application/x-www-form-urlencoded";
      (reqOptions.headers as any)["Content-Length"] = Buffer.byteLength(body);
    }

    const req = https.request(reqOptions, (res) => {
      // Parse and save any cookies from response
      const setCookies = parseSetCookies(res.headers["set-cookie"] as any);
      if (Object.keys(setCookies).length > 0) {
        saveCookies(setCookies);
      }

      // Handle redirects
      if (
        followRedirects &&
        maxRedirects > 0 &&
        res.statusCode &&
        [301, 302, 303, 307, 308].includes(res.statusCode) &&
        res.headers.location
      ) {
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

      const chunks: Buffer[] = [];
      res.on("data", (chunk: Buffer) => {
        chunks.push(chunk);
      });
      res.on("end", () => {
        const body = Buffer.concat(chunks).toString("utf-8");
        resolve({
          statusCode: res.statusCode || 0,
          headers: res.headers as Record<string, any>,
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
