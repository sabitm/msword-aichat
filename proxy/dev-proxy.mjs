#!/usr/bin/env node
import http from "node:http";
import https from "node:https";
import { URL } from "node:url";

const PORT = Number(process.env.PROXY_PORT ?? 8787);
const TARGET = process.env.TARGET_URL ?? "https://api.openai.com";
const targetUrl = new URL(TARGET);

function setCorsHeaders(response) {
  response.setHeader("Access-Control-Allow-Origin", "*");
  response.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  response.setHeader(
    "Access-Control-Allow-Headers",
    "Authorization, Content-Type, x-api-key, anthropic-version, Accept",
  );
}

const server = http.createServer((request, response) => {
  setCorsHeaders(response);

  if (request.method === "OPTIONS") {
    response.writeHead(204);
    response.end();
    return;
  }

  const upstreamPath = request.url ?? "/";
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
  console.log("Set base URL to http://localhost:8787/v1 in add-in settings.");
});