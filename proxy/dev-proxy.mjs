#!/usr/bin/env node
import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { URL } from "node:url";

const PORT = Number(process.env.PROXY_PORT ?? 8787);
const TARGET = process.env.TARGET_URL ?? "https://api.openai.com";
const targetUrl = new URL(TARGET);
const LOG_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "logs");
const DEBUG_LOG_FILE = path.join(LOG_DIR, "agent-debug.jsonl");

function setCorsHeaders(response) {
  response.setHeader("Access-Control-Allow-Origin", "*");
  response.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  response.setHeader(
    "Access-Control-Allow-Headers",
    "Authorization, Content-Type, x-api-key, anthropic-version, Accept",
  );
}

function appendDebugLog(body) {
  fs.mkdirSync(LOG_DIR, { recursive: true });
  const line = body.trim();
  if (!line) {
    return;
  }
  fs.appendFileSync(DEBUG_LOG_FILE, `${line}\n`, "utf8");
}

function readRequestBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";
    request.on("data", (chunk) => {
      body += chunk;
    });
    request.on("end", () => resolve(body));
    request.on("error", reject);
  });
}

function handleDebugLog(request, response) {
  readRequestBody(request)
    .then((body) => {
      appendDebugLog(body);
      setCorsHeaders(response);
      response.writeHead(204);
      response.end();
    })
    .catch((error) => {
      setCorsHeaders(response);
      response.writeHead(500, { "Content-Type": "application/json" });
      response.end(JSON.stringify({ error: error.message }));
    });
}

const server = http.createServer((request, response) => {
  setCorsHeaders(response);

  if (request.method === "OPTIONS") {
    response.writeHead(204);
    response.end();
    return;
  }

  const upstreamPath = request.url ?? "/";
  const parsedPath = new URL(upstreamPath, "http://localhost");

  if (parsedPath.pathname === "/__debug/log") {
    if (request.method === "POST") {
      handleDebugLog(request, response);
      return;
    }

    if (request.method === "GET") {
      response.writeHead(200, { "Content-Type": "application/json" });
      response.end(
        JSON.stringify({
          ok: true,
          file: DEBUG_LOG_FILE,
          message: "POST JSONL lines to append agent debug events.",
        }),
      );
      return;
    }
  }

  const transport = targetUrl.protocol === "https:" ? https : http;

  const upstream = transport.request(
    {
      protocol: targetUrl.protocol,
      hostname: targetUrl.hostname,
      port: targetUrl.port || undefined,
      path: `${targetUrl.pathname.replace(/\/$/, "")}${upstreamPath}`,
      method: request.method,
      headers: {
        ...request.headers,
        host: targetUrl.host,
      },
    },
    (upstreamResponse) => {
      setCorsHeaders(response);
      response.writeHead(upstreamResponse.statusCode ?? 502, upstreamResponse.headers);
      upstreamResponse.pipe(response);
    },
  );

  upstream.on("error", (error) => {
    setCorsHeaders(response);
    response.writeHead(502, { "Content-Type": "application/json" });
    response.end(JSON.stringify({ error: error.message }));
  });

  request.pipe(upstream);
});

server.listen(PORT, () => {
  console.log(`CORS proxy listening on http://localhost:${PORT}`);
  console.log(`Forwarding to ${TARGET}`);
  console.log(`Agent debug log file: ${DEBUG_LOG_FILE}`);
  console.log("Set base URL to http://localhost:8787/v1 in add-in settings.");
});