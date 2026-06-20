#!/usr/bin/env node
import { execSync, spawn } from "node:child_process";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import https from "node:https";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const baseUrl = process.env.SMOKE_BASE_URL ?? "https://localhost:3000";

const results = [];

function pass(name, detail = "") {
  results.push({ name, ok: true, detail });
  console.log(`  ✓ ${name}${detail ? ` — ${detail}` : ""}`);
}

function fail(name, detail = "") {
  results.push({ name, ok: false, detail });
  console.error(`  ✗ ${name}${detail ? ` — ${detail}` : ""}`);
}

function run(cmd, name) {
  try {
    execSync(cmd, { cwd: root, stdio: "pipe" });
    pass(name);
    return true;
  } catch (error) {
    const message = error.stderr?.toString().trim() || error.message;
    fail(name, message.split("\n").pop());
    return false;
  }
}

function fetchText(url) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, { rejectUnauthorized: false }, (response) => {
      let body = "";
      response.on("data", (chunk) => {
        body += chunk;
      });
      response.on("end", () => {
        resolve({ status: response.statusCode ?? 0, body, headers: response.headers });
      });
    });
    request.on("error", reject);
    request.setTimeout(10_000, () => {
      request.destroy(new Error("timeout"));
    });
  });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForDevServer(maxAttempts = 30) {
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    try {
      const { status } = await fetchText(`${baseUrl}/taskpane.html`);
      if (status === 200) return true;
    } catch {
      // retry
    }
    await sleep(1000);
  }
  return false;
}

function checkDistBundle() {
  const dist = join(root, "dist");
  if (!existsSync(dist)) {
    fail("dist/ exists");
    return;
  }
  pass("dist/ exists");

  for (const file of ["taskpane.html", "commands.html"]) {
    if (existsSync(join(dist, file))) pass(`dist/${file}`);
    else fail(`dist/${file}`);
  }

  const distFiles = readdirSync(dist);
  const bundle = distFiles.find((name) => name.endsWith(".bundle.js"));
  if (bundle) {
    const size = readFileSync(join(dist, bundle)).length;
    pass("dist webpack bundle", `${bundle} (${Math.round(size / 1024)} KB)`);
    if (size < 100_000) fail("dist bundle size", `${size} bytes — suspiciously small`);
    else if (size > 3_000_000) fail("dist bundle size", `${size} bytes — exceeds 3 MB budget`);
    else pass("dist bundle size within budget");
  } else {
    fail("dist webpack bundle");
  }

  const assetsDir = join(dist, "assets");
  if (!existsSync(assetsDir)) {
    fail("dist/assets/");
    return;
  }
  pass("dist/assets/");

  const assets = readdirSync(assetsDir);
  const iconCount = assets.filter((name) => name.endsWith(".png")).length;
  if (iconCount >= 3) pass("dist icon assets", `${iconCount} PNG files`);
  else fail("dist icon assets", `found ${iconCount}`);

  const taskpaneHtml = readFileSync(join(dist, "taskpane.html"), "utf8");
  if (taskpaneHtml.includes(".bundle.js")) pass("taskpane.html references bundle");
  else fail("taskpane.html references bundle");
  if (!taskpaneHtml.includes('type="module"')) pass("taskpane.html no ESM module script");
  else fail("taskpane.html no ESM module script");
}

function checkPackageOutput() {
  const packageDir = join(root, "package");
  if (!existsSync(packageDir)) {
    fail("package/ output");
    return;
  }
  pass("package/ output");

  const manifest = readFileSync(join(packageDir, "manifest.xml"), "utf8");
  if (manifest.includes("https://YOUR_DEPLOY_HOST")) {
    fail("package manifest URLs substituted");
  } else {
    pass("package manifest URLs substituted");
  }

  for (const file of ["taskpane.html", "commands.html", "README.txt"]) {
    if (existsSync(join(packageDir, file))) pass(`package/${file}`);
    else fail(`package/${file}`);
  }

  const packageFiles = readdirSync(packageDir);
  const bundle = packageFiles.find((name) => name.endsWith(".bundle.js"));
  if (bundle) pass("package webpack bundle", bundle);
  else fail("package webpack bundle");
}

function checkToolRegistry() {
  const source = readFileSync(join(root, "src/agent/tools/registry.ts"), "utf8");
  const expectedTools = [
    "get_selection",
    "get_document_text",
    "search_document",
    "insert_text",
    "replace_text",
    "find_and_replace",
    "replace_at_match",
    "delete_range",
    "apply_style",
    "format_range",
    "insert_comment",
    "insert_table",
    "list_tables",
    "update_table",
  ];

  for (const tool of expectedTools) {
    if (source.includes(`name: "${tool}"`)) pass(`tool registered: ${tool}`);
    else fail(`tool registered: ${tool}`);
  }
}

function checkSlashCommands() {
  const source = readFileSync(join(root, "src/agent/slash-commands.ts"), "utf8");
  for (const command of ["fix", "table", "toc", "summarize", "formal", "comment"]) {
    if (source.includes(`name: "${command}"`)) pass(`slash command: /${command}`);
    else fail(`slash command: /${command}`);
  }
}

function checkNoViteStack() {
  const removed = [
    "vite.config.ts",
    "src/taskpane/main.tsx",
    "src/taskpane/App.tsx",
    "src/settings/store.ts",
  ];
  for (const file of removed) {
    if (!existsSync(join(root, file))) pass(`removed: ${file}`);
    else fail(`removed: ${file}`);
  }
}

async function checkDevServer() {
  const endpoints = [
    "/taskpane.html",
    "/commands.html",
    "/assets/icon-32.png",
  ];

  for (const path of endpoints) {
    try {
      const { status, body } = await fetchText(`${baseUrl}${path}`);
      if (status !== 200) {
        fail(`GET ${path}`, `HTTP ${status}`);
        continue;
      }
      if (path.endsWith(".html") && !body.includes("<html")) {
        fail(`GET ${path}`, "missing html root");
        continue;
      }
      pass(`GET ${path}`, `HTTP ${status}`);
    } catch (error) {
      fail(`GET ${path}`, error instanceof Error ? error.message : "request failed");
    }
  }

  try {
    const { body } = await fetchText(`${baseUrl}/taskpane.html`);
    const bundleEntry = body.match(/src="(\/taskpane[^"]*\.bundle\.js)"/);
    const legacyViteEntry = body.match(/src="(\/src\/taskpane\/main\.tsx[^"]*)"/);

    if (legacyViteEntry) {
      fail("taskpane entry script reference", "still using Vite ESM entry");
      return;
    }

    const entryPath = bundleEntry?.[1];
    if (!entryPath) {
      fail("taskpane entry script reference");
      return;
    }

    pass("taskpane entry script reference", entryPath);

    const { status, body: js } = await fetchText(`${baseUrl}${entryPath}`);
    if (status === 200 && js.length >= 100_000) {
      pass("taskpane bundle served", `HTTP ${status}, ${Math.round(js.length / 1024)} KB`);
    } else {
      fail("taskpane bundle served", `HTTP ${status}, ${js.length} bytes`);
    }
  } catch (error) {
    fail("taskpane bundle served", error instanceof Error ? error.message : "request failed");
  }
}

function checkProxy() {
  const proxyPath = join(root, "proxy/dev-proxy.mjs");
  if (existsSync(proxyPath)) pass("proxy/dev-proxy.mjs exists");
  else fail("proxy/dev-proxy.mjs exists");

  try {
    execSync("node --check proxy/dev-proxy.mjs", { cwd: root, stdio: "pipe" });
    pass("proxy script syntax valid");
  } catch (error) {
    fail("proxy script syntax valid", error.message);
  }
}

console.log("\n=== msword-aichat smoke test ===\n");

console.log("Build & validate");
run("npm run typecheck", "npm run typecheck");
run("npm run build", "npm run build");
run("npm run validate", "npm run validate");
run("npm run validate:prod", "npm run validate:prod");
run('npm run package -- https://addins.example.com/msword-aichat', "npm run package");

console.log("\nArtifacts");
checkDistBundle();
checkPackageOutput();

console.log("\nSource contracts");
checkToolRegistry();
checkSlashCommands();
checkNoViteStack();

console.log("\nDev server (" + baseUrl + ")");
let devProcess = null;
const skipDevServer = process.env.SMOKE_SKIP_DEV_SERVER === "1";

if (skipDevServer) {
  console.log("  (skipped — SMOKE_SKIP_DEV_SERVER=1)");
} else {
  devProcess = spawn("npm", ["run", "dev"], {
    cwd: root,
    stdio: "ignore",
    shell: true,
    detached: process.platform !== "win32",
  });

  const ready = await waitForDevServer();
  if (!ready) {
    fail("dev server ready");
  } else {
    pass("dev server ready");
    await checkDevServer();
  }
}

console.log("\nProxy");
checkProxy();

if (devProcess?.pid) {
  try {
    if (process.platform === "win32") {
      execSync(`taskkill /PID ${devProcess.pid} /T /F`, { stdio: "ignore" });
    } else {
      process.kill(-devProcess.pid);
    }
  } catch {
    // best-effort cleanup
  }
}

const failed = results.filter((result) => !result.ok);
const passed = results.filter((result) => result.ok);

console.log("\n=== Summary ===");
console.log(`Passed: ${passed.length}`);
console.log(`Failed: ${failed.length}`);

if (failed.length) {
  console.log("\nFailures:");
  for (const item of failed) {
    console.log(`  - ${item.name}${item.detail ? `: ${item.detail}` : ""}`);
  }
  process.exit(1);
}

console.log("\nManual Word 2016 checks (not run in automation):");
console.log("  - npm start → Home → AI Chat → pane loads (not blank)");
console.log("  - Settings save, test connection, fetch models");
console.log("  - Chat stream, agent replace_text Apply/Undo, slash /fix");
console.log("  - See README Word 2016 QA matrix for full sign-off");
console.log("\nAll automated smoke checks passed.\n");