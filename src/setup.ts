import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const envFile = path.join(root, ".env");
const exampleFile = path.join(root, ".env.example");

console.log("🛠️  Setting up h-office MCP...");

if (!fs.existsSync(envFile) && fs.existsSync(exampleFile)) {
  fs.copyFileSync(exampleFile, envFile);
  console.log("✅ Created .env file from .env.example");
  console.log("👉 Please open .env file and set your H_OFFICE_EMAIL and H_OFFICE_PASSWORD.");
} else if (fs.existsSync(envFile)) {
  console.log("✅ .env file already exists.");
}

console.log("🎉 Setup script completed.");
