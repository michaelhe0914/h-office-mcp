import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.join(path.dirname(fileURLToPath(import.meta.url)), ".."));
const envFile = path.join(root, ".env");
const exampleFile = path.join(root, ".env.example");
const distIndex = path.join(root, "dist", "index.js").replace(/\\/g, "/");

console.log("==================================================");
console.log("🛠️  金安出書統計表 h-office MCP 伺服器初始化設定");
console.log("==================================================");

// 1. Check / Create .env
if (!fs.existsSync(envFile) && fs.existsSync(exampleFile)) {
  fs.copyFileSync(exampleFile, envFile);
  console.log("✅ 已由 .env.example 建立 .env 設定檔。");
} else if (fs.existsSync(envFile)) {
  console.log("✅ .env 設定檔已存在。");
}

// 2. Generate sample Antigravity config snippet
const antigravitySnippet = {
  mcpServers: {
    "h-office": {
      command: "node",
      args: [distIndex],
      env: {
        H_OFFICE_EMAIL: "your_email@gmail.com",
        H_OFFICE_PASSWORD: "your_password",
      },
    },
  },
};

const configFile = path.join(root, "antigravity_mcp_config.json");
fs.writeFileSync(configFile, JSON.stringify(antigravitySnippet, null, 2), "utf-8");
console.log(`✅ 已為您產生本地專用 Antigravity MCP 設定範本：`);
console.log(`   ${configFile}`);

console.log("\n👉 下一步操作指引：");
console.log("1. 請用文字編輯器開啟 `.env` 檔案，填入您個人的 h-office 帳號 (Email) 與密碼。");
console.log("2. 在 Antigravity 的 MCP 設定中加入以下設定片段：\n");
console.log(JSON.stringify(antigravitySnippet, null, 2));
console.log("\n3. 若您的 Google 帳號有雙重驗證 (2FA)，可先執行 `npm run login` 完成首次登入。");
console.log("==================================================\n");
