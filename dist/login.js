/**
 * Login helper for h-office using Puppeteer to automate Google OAuth.
 *
 * Usage: node dist/login.js
 *
 * This script will:
 * 1. Open a browser window
 * 2. Navigate to h-office login page
 * 3. Automate Google OAuth sign-in
 * 4. Capture and save the session cookie
 */
import puppeteer from "puppeteer";
import { saveCookies, getCookiesFilePath } from "./client.js";
const BASE_URL = process.env.H_OFFICE_BASE_URL || "https://h-office.king-an.com.tw:8082";
const EMAIL = process.env.H_OFFICE_EMAIL || "";
const PASSWORD = process.env.H_OFFICE_PASSWORD || "";
async function login() {
    console.log("🔐 Starting h-office login via Google OAuth...");
    console.log(`   Target: ${BASE_URL}`);
    if (!EMAIL || !PASSWORD) {
        console.error("❌ Please set H_OFFICE_EMAIL and H_OFFICE_PASSWORD environment variables.");
        process.exit(1);
    }
    const browser = await puppeteer.launch({
        headless: false, // Show browser for potential 2FA / CAPTCHA
        args: [
            "--ignore-certificate-errors",
            "--no-sandbox",
            "--disable-setuid-sandbox",
        ],
        // Use larger viewport to avoid mobile layout
        defaultViewport: { width: 1280, height: 800 },
    });
    try {
        const page = await browser.newPage();
        // Ignore SSL errors
        await page.setBypassCSP(true);
        console.log("📡 Navigating to login page...");
        await page.goto(`${BASE_URL}/login`, {
            waitUntil: "networkidle2",
            timeout: 30000,
        });
        // We should now be on Google's sign-in page
        console.log("📧 Entering email...");
        // Wait for email input
        await page.waitForSelector('input[type="email"]', { timeout: 15000 });
        await page.type('input[type="email"]', EMAIL, { delay: 50 });
        // Click Next
        const nextButtons = await page.$$('button');
        for (const btn of nextButtons) {
            const text = await page.evaluate(el => el.textContent, btn);
            if (text && (text.includes('Next') || text.includes('下一步') || text.includes('繼續'))) {
                await btn.click();
                break;
            }
        }
        // Wait for password input
        console.log("🔑 Entering password...");
        await page.waitForSelector('input[type="password"]', { visible: true, timeout: 15000 });
        await new Promise(r => setTimeout(r, 1000)); // Small delay for animation
        await page.type('input[type="password"]', PASSWORD, { delay: 50 });
        // Click Next again
        await new Promise(r => setTimeout(r, 500));
        const nextButtons2 = await page.$$('button');
        for (const btn of nextButtons2) {
            const text = await page.evaluate(el => el.textContent, btn);
            if (text && (text.includes('Next') || text.includes('下一步') || text.includes('繼續'))) {
                await btn.click();
                break;
            }
        }
        // Wait for redirect back to h-office
        console.log("⏳ Waiting for OAuth redirect back to h-office...");
        console.log("   (If you see a consent screen or 2FA prompt, please complete it manually)");
        try {
            await page.waitForFunction((baseUrl) => window.location.href.startsWith(baseUrl), { timeout: 60000 }, BASE_URL);
        }
        catch {
            // Check if we're still on Google (might need user interaction)
            const currentUrl = page.url();
            if (currentUrl.includes("accounts.google.com")) {
                console.log("⚠️  Still on Google page. Please complete the login manually in the browser.");
                console.log("   Waiting up to 120 seconds...");
                await page.waitForFunction((baseUrl) => window.location.href.startsWith(baseUrl), { timeout: 120000 }, BASE_URL);
            }
        }
        console.log("✅ Redirected back to h-office!");
        // Extract cookies
        const cookies = await page.cookies();
        const cookieMap = {};
        for (const cookie of cookies) {
            if (cookie.domain.includes("king-an.com.tw")) {
                cookieMap[cookie.name] = cookie.value;
            }
        }
        if (Object.keys(cookieMap).length > 0) {
            saveCookies(cookieMap);
            console.log(`✅ Session cookie saved to: ${getCookiesFilePath()}`);
            console.log(`   Cookies: ${Object.keys(cookieMap).join(", ")}`);
            // Verify login
            const pageContent = await page.content();
            if (!pageContent.includes("尚未登入")) {
                console.log("✅ Login verified successfully!");
            }
            else {
                console.log("⚠️  Cookie saved but login status unclear. The cookie might still work.");
            }
        }
        else {
            console.error("❌ No cookies captured. Login may have failed.");
        }
    }
    catch (error) {
        console.error("❌ Login failed:", error);
    }
    finally {
        await browser.close();
    }
}
login().catch(console.error);
