#!/usr/bin/env node
import { cpSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const distDir = join(root, "dist");
const packageDir = join(root, "package");
const deployHost = process.argv[2]?.replace(/\/+$/, "");

if (!existsSync(distDir)) {
  console.error("dist/ not found. Run npm run build first.");
  process.exit(1);
}

mkdirSync(packageDir, { recursive: true });

for (const entry of ["assets", "taskpane.html", "commands.html"]) {
  const source = join(distDir, entry);
  const target = join(packageDir, entry);
  if (!existsSync(source)) {
    console.error(`Missing build artifact: ${source}`);
    process.exit(1);
  }
  cpSync(source, target, { recursive: true });
}

let manifest = readFileSync(join(root, "manifest.prod.xml"), "utf8");
if (deployHost) {
  manifest = manifest.replaceAll("https://YOUR_DEPLOY_HOST", deployHost);
  console.log(`Using deploy host: ${deployHost}`);
} else {
  console.warn("No deploy host passed. manifest.xml still contains YOUR_DEPLOY_HOST placeholders.");
  console.warn("Usage: npm run package -- https://addins.example.com/msword-aichat");
}

writeFileSync(join(packageDir, "manifest.xml"), manifest);

const readme = `Word AI Chat — distribution package

Contents:
- manifest.xml  (production manifest)
- taskpane.html, commands.html
- assets/         (icons and bundled JS/CSS)

Deploy:
1. Upload all files in this folder to your HTTPS origin.
2. Ensure CORS allows the add-in origin if calling remote LLM gateways.
3. Sideload manifest.xml for testing, or publish via Microsoft 365 admin center / AppSource.

Regenerate with a host URL:
  npm run package -- https://your-host.example.com/path
`;

writeFileSync(join(packageDir, "README.txt"), readme);
console.log(`Package written to ${packageDir}`);