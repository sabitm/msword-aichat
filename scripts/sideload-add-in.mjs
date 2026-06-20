#!/usr/bin/env node
import { createRequire } from "node:module";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { getRuntimeRoot } from "./runtime-root.mjs";

var rootDir = getRuntimeRoot();
var require = createRequire(join(rootDir, "package.json"));

async function main() {
  var manifestPath = join(rootDir, "manifest.xml");
  if (!existsSync(manifestPath)) {
    throw new Error("manifest.xml not found next to the launcher: " + manifestPath);
  }

  console.log("");
  console.log("Word AI Chat — register add-in with Word");
  console.log("Manifest: " + manifestPath);
  console.log("");

  var devSettings = require("office-addin-dev-settings");
  await devSettings.registerAddIn(manifestPath);

  console.log("Add-in registered.");
  console.log("Next:");
  console.log("  1. Run Launch-server.bat (or msword-aichat-server.exe)");
  console.log("  2. Open Word -> Home -> AI Chat");
  console.log("");
}

main().catch(function (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});