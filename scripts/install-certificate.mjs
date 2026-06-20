#!/usr/bin/env node
import { createRequire } from "node:module";
import { join } from "node:path";
import { getRuntimeRoot } from "./runtime-root.mjs";

var rootDir = getRuntimeRoot();
var require = createRequire(join(rootDir, "package.json"));

async function main() {
  console.log("");
  console.log("Word AI Chat — install localhost HTTPS certificate");
  console.log("Run as Administrator if Windows prompts for permission.");
  console.log("");

  var devCerts = require("office-addin-dev-certs");
  await devCerts.ensureCertificatesAreInstalled(undefined, undefined, true);

  console.log("");
  console.log("Certificate ready for https://localhost:3000");
  console.log("");
}

main().catch(function (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});