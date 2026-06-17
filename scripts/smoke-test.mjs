#!/usr/bin/env node
import { execSync } from "node:child_process";
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

  const assetsDir = join(dist, "assets");
  if (!existsSync(assetsDir)) {
    fail("dist/assets/");
    return;
  }

  const assets = readdirSync(assetsDir);
  const hasJs = assets.some((name) => name.endsWith(".js"));
  const hasCss = assets.some((name) => name.endsWith(".css"));
  const hasFluentChunk = assets.some((name) => name.includes("fluent"));
  const hasReactChunk = assets.some((name) => name.includes("react"));

  if (hasJs) pass("dist JS bundle");
  else fail("dist JS bundle");
  if (hasCss) pass("dist CSS bundle");
  else fail("dist CSS bundle");
  if (hasFluentChunk) pass("fluent code-split chunk");
  else fail("fluent code-split chunk");
  if (hasReactChunk) pass("react code-split chunk");
  else fail("react code-split chunk");
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
}

function checkToolRegistry() {
  const source = readFileSync(join(root, "src/agent/tools/registry.ts"), "utf8");
  const expectedTools = [
    "get_selection",
    "get_document_text",
    "search_document",
    "insert_text",
    "replace_text",
    "delete_range",
    "apply_style",
    "format_range",
    "insert_comment",
    "insert_table",
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
      if (path.endsWith(".html") && path.includes("taskpane") && !body.includes("taskpane")) {
        fail(`GET ${path}`, "missing taskpane entry script");
        continue;
      }
      pass(`GET ${path}`, `HTTP ${status}`);
    } catch (error) {
      fail(`GET ${path}`, error instanceof Error ? error.message : "request failed");
    }
  }

  try {
    const { body } = await fetchText(`${baseUrl}/taskpane.html`);
    const prodEntry = body.match(/src="(\/assets\/[^"]+\.js)"/);
    const devEntry = body.match(/src="(\/src\/taskpane\/main\.tsx[^"]*)"/);
    const entryPath = prodEntry?.[1] ?? devEntry?.[1];

    if (!entryPath) {
      fail("taskpane entry script reference");
      return;
    }

    pass("taskpane entry script reference", entryPath);

    const entryUrl = entryPath.split("?")[0];
    const { status, body: js } = await fetchText(`${baseUrl}${entryPath}`);
    const isDevEntry = entryUrl.includes("/src/");
    const minSize = isDevEntry ? 500 : 10_000;
    if (status === 200 && js.length >= minSize) {
      pass("taskpane entry served", `${isDevEntry ? "dev" : "prod"} HTTP ${status}`);
    } else {
      fail("taskpane entry served", `HTTP ${status}, ${js.length} bytes`);
    }
  } catch (error) {
    fail("taskpane entry served", error instanceof Error ? error.message : "request failed");
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

console.log("\nDev server (" + baseUrl + ")");
await checkDevServer();

console.log("\nProxy");
checkProxy();

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

console.log("\nManual Word host checks (not run in CI/Linux):");
console.log("  - Word Desktop: npm start → Home → AI Chat → onboarding → chat/agent");
console.log("  - Word on the web: upload package/ manifest → run QA matrix in README");
console.log("  - Agent tools: get_selection, replace_text Apply/Undo, insert_comment");
console.log("\nAll automated smoke checks passed.\n");