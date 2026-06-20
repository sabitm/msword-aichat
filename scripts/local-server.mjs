#!/usr/bin/env node
import { createServer } from "node:https";
import { createReadStream, existsSync, readFileSync, statSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { homedir } from "node:os";
import { getRuntimeRoot } from "./runtime-root.mjs";

var PORT = Number(process.env.MSWORD_AICHAT_PORT || 3000);
var runtimeRoot = getRuntimeRoot();
var appRoot = process.env.MSWORD_AICHAT_APP_DIR || join(runtimeRoot, "app");
if (!existsSync(appRoot)) {
  var distFallback = join(runtimeRoot, "..", "dist");
  if (existsSync(distFallback)) {
    appRoot = distFallback;
  }
}

var MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".map": "application/json; charset=utf-8",
};

function resolveContentType(filePath) {
  return MIME_TYPES[extname(filePath).toLowerCase()] || "application/octet-stream";
}

function resolveFilePath(urlPath) {
  var decoded = decodeURIComponent(urlPath.split("?")[0]);
  if (decoded === "/" || decoded === "") {
    decoded = "/taskpane.html";
  }
  var filePath = normalize(join(appRoot, decoded));
  if (!filePath.startsWith(normalize(appRoot))) {
    return null;
  }
  return filePath;
}

async function getHttpsOptions() {
  try {
    var modulePath = await import("node:module");
    var require = modulePath.createRequire(join(runtimeRoot, "package.json"));
    var devCerts = require("office-addin-dev-certs");
    return await devCerts.getHttpsServerOptions();
  } catch (_moduleError) {
    var certDir = process.env.OFFICE_ADDIN_CERT_DIR || join(homedir(), ".office-addin-dev-certs");
    var keyPath = join(certDir, "localhost.key");
    var certPath = join(certDir, "localhost.crt");
    if (existsSync(keyPath) && existsSync(certPath)) {
      return {
        key: readFileSync(keyPath),
        cert: readFileSync(certPath),
      };
    }
    throw new Error(
      "HTTPS certificates not found. Run Install-certificate.bat as Administrator once, then retry.",
    );
  }
}

function sendFile(filePath, res) {
  var stream = createReadStream(filePath);
  stream.on("error", function () {
    res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Failed to read file.");
  });
  res.writeHead(200, {
    "Content-Type": resolveContentType(filePath),
    "Access-Control-Allow-Origin": "*",
  });
  stream.pipe(res);
}

async function main() {
  if (!existsSync(appRoot)) {
    console.error("App folder not found:", appRoot);
    process.exit(1);
  }

  var httpsOptions = await getHttpsOptions();
  var server = createServer(function (req, res) {
    if (req.method === "OPTIONS") {
      res.writeHead(204, {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, HEAD, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization",
      });
      res.end();
      return;
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      res.writeHead(405, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Method not allowed");
      return;
    }

    var filePath = resolveFilePath(req.url || "/");
    if (!filePath || !existsSync(filePath) || !statSync(filePath).isFile()) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not found");
      return;
    }

    if (req.method === "HEAD") {
      res.writeHead(200, {
        "Content-Type": resolveContentType(filePath),
        "Access-Control-Allow-Origin": "*",
      });
      res.end();
      return;
    }

    sendFile(filePath, res);
  });

  server.listen(PORT, "localhost", function () {
    console.log("");
    console.log("Word AI Chat local server running");
    console.log("  https://localhost:" + PORT + "/taskpane.html");
    console.log("");
    console.log("Keep this window open while using Word.");
    console.log("Press Ctrl+C to stop.");
    console.log("");
  });
}

main().catch(function (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});