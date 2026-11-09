#!/usr/bin/env node
/*
 * Preview the launcher on your own computer, with the same security headers Azure will send.
 *
 * Usage (run from the project folder):
 *   node scripts/preview-local.mjs            (uses port 8080)
 *   node scripts/preview-local.mjs --port 9000
 *
 * Then open http://localhost:8080 in Edge or Chrome. Press Ctrl+C to stop.
 * Outside Teams, the badge at the bottom should change to "Web browser" within a few seconds.
 */
import { createServer } from "node:http";
import { readFileSync, existsSync, statSync } from "node:fs";
import { dirname, extname, join, normalize, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const webRoot = join(root, "web");
const config = JSON.parse(readFileSync(join(webRoot, "staticwebapp.config.json"), "utf8"));
const portIndex = process.argv.indexOf("--port");
const port = portIndex > -1 ? Number(process.argv[portIndex + 1]) : 8080;

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon",
  ".txt": "text/plain; charset=utf-8"
};

function headersFor(route) {
  const headers = { ...(config.globalHeaders || {}) };
  delete headers["Strict-Transport-Security"]; // only meaningful over HTTPS
  for (const rule of config.routes || []) {
    if (rule.route === route && rule.headers) Object.assign(headers, rule.headers);
  }
  return headers;
}

function resolveFile(urlPath) {
  let clean;
  try {
    clean = decodeURIComponent(urlPath.split("?")[0]);
  } catch (error) {
    return null;
  }
  const relative = clean === "/" ? "index.html" : clean.replace(/^\/+/, "");
  const full = normalize(join(webRoot, relative));
  if (!full.startsWith(webRoot + sep)) return null;
  if (relative === "staticwebapp.config.json") return null;
  if (!existsSync(full) || !statSync(full).isFile()) return null;
  return full;
}

const server = createServer((request, response) => {
  const route = request.url.split("?")[0];
  const file = resolveFile(request.url);
  if (!file) {
    response.writeHead(404, { ...headersFor(route), "Content-Type": TYPES[".html"] });
    response.end(readFileSync(join(webRoot, "404.html")));
    return;
  }
  response.writeHead(200, { ...headersFor(route), "Content-Type": TYPES[extname(file)] || "application/octet-stream" });
  response.end(readFileSync(file));
});

server.listen(port, "127.0.0.1", () => {
  console.log(`\nPreview running at http://localhost:${port}`);
  console.log("Same security headers as Azure (from web/staticwebapp.config.json). Press Ctrl+C to stop.\n");
});
