#!/usr/bin/env node
"use strict";
/*
 * Job Tracker — local server (no dependencies, Node 18+).
 *   - serves the app from ./public
 *   - /api/store    saves your data to ./data/tracker.json
 *   - /api/claude   tailors resumes/cover letters with the Claude API (your key stays here)
 *   - /api/jobs     searches jobs through the Adzuna API
 * It listens on 127.0.0.1 only, so nothing outside your computer can reach it.
 */
const http = require("http");
const fs = require("fs");
const path = require("path");
const { loadEnv, config } = require("./lib/env");
const { Store } = require("./lib/store");
const { askClaude } = require("./lib/anthropic");
const { searchJobs, COUNTRIES } = require("./lib/adzuna");

loadEnv();

const PUBLIC = path.join(__dirname, "public");
const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".json": "application/json", ".svg": "image/svg+xml", ".png": "image/png", ".ico": "image/x-icon", ".woff2": "font/woff2" };
const MAX_BODY = 2 * 1024 * 1024;

function createServer(cfg = config(), store = new Store(cfg.dataDir)) {
  const send = (res, status, obj) => {
    const body = JSON.stringify(obj);
    res.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
    res.end(body);
  };
  const sendErr = (res, e) => send(res, e.status || 500, { error: { code: e.code || "server_error", message: e.message || "Server error" } });

  const readBody = (req) => new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on("data", (c) => {
      size += c.length;
      if (size > MAX_BODY) { reject(Object.assign(new Error("Request too large"), { status: 413, code: "too_large" })); req.destroy(); return; }
      chunks.push(c);
    });
    req.on("end", () => {
      try { resolve(chunks.length ? JSON.parse(Buffer.concat(chunks).toString("utf8")) : {}); }
      catch { reject(Object.assign(new Error("Body must be JSON"), { status: 400, code: "bad_request" })); }
    });
    req.on("error", reject);
  });

  // Only accept API calls from this app's own pages (blocks other websites in your browser from using your keys).
  const sameOrigin = (req) => {
    const origin = req.headers.origin;
    if (!origin) return true;
    try { const u = new URL(origin); return ["127.0.0.1", "localhost", "[::1]"].includes(u.hostname) && u.host === req.headers.host; } catch { return false; }
  };

  async function api(req, res, url) {
    if (!sameOrigin(req)) return send(res, 403, { error: { code: "forbidden", message: "Cross-origin request blocked" } });
    const route = url.pathname;

    if (route === "/api/status" && req.method === "GET") {
      return send(res, 200, {
        claude: !!cfg.anthropicKey, model: cfg.model, fastModel: cfg.fastModel,
        jobs: !!(cfg.adzunaId && cfg.adzunaKey), countries: COUNTRIES,
        dataFile: path.relative(process.cwd(), store.file) || store.file,
      });
    }

    if (route === "/api/store") {
      if (req.method === "GET") return send(res, 200, store.all());
      const p = url.searchParams.get("path") || "";
      if (req.method === "PUT") { await store.set(p, await readBody(req)); return send(res, 200, { ok: true }); }
      if (req.method === "DELETE") { await store.delete(p); return send(res, 200, { ok: true }); }
    }

    if (route === "/api/claude" && req.method === "POST") {
      const body = await readBody(req);
      const ctl = new AbortController();
      res.on("close", () => { if (!res.writableEnded) ctl.abort(); });
      const timer = setTimeout(() => ctl.abort(), 5 * 60 * 1000);
      try { return send(res, 200, await askClaude(cfg, body, ctl.signal)); }
      finally { clearTimeout(timer); }
    }

    if (route === "/api/jobs/search" && req.method === "GET") {
      const q = Object.fromEntries(url.searchParams.entries());
      return send(res, 200, { results: await searchJobs(cfg, q) });
    }

    return send(res, 404, { error: { code: "not_found", message: "Unknown API route" } });
  }

  function serveStatic(req, res, url) {
    let rel = decodeURIComponent(url.pathname);
    if (rel === "/") rel = "/index.html";
    const file = path.normalize(path.join(PUBLIC, rel));
    if (!file.startsWith(PUBLIC + path.sep)) { res.writeHead(403); return res.end("Forbidden"); }
    fs.stat(file, (err, st) => {
      if (err || !st.isFile()) { res.writeHead(404, { "content-type": "text/plain" }); return res.end("Not found"); }
      res.writeHead(200, {
        "content-type": TYPES[path.extname(file)] || "application/octet-stream",
        "cache-control": rel.startsWith("/vendor/") ? "public, max-age=86400" : "no-cache",
        "x-content-type-options": "nosniff",
      });
      fs.createReadStream(file).pipe(res);
    });
  }

  return http.createServer(async (req, res) => {
    const url = new URL(req.url, "http://localhost");
    try {
      if (url.pathname.startsWith("/api/")) await api(req, res, url);
      else if (req.method === "GET" || req.method === "HEAD") serveStatic(req, res, url);
      else { res.writeHead(405); res.end(); }
    } catch (e) {
      if (!e.status || e.status >= 500) console.error(e);
      if (!res.headersSent) sendErr(res, e);
    }
  });
}

if (require.main === module) {
  const cfg = config();
  const server = createServer(cfg);
  server.on("error", (e) => {
    if (e.code === "EADDRINUSE") console.error(`\nPort ${cfg.port} is already in use. Set PORT=3001 in .env (or stop the other app) and try again.\n`);
    else console.error(e);
    process.exit(1);
  });
  server.listen(cfg.port, cfg.host, () => {
    const url = `http://localhost:${cfg.port}`;
    console.log(`\n  Job Tracker is running → ${url}\n`);
    console.log(`  Claude tailoring : ${cfg.anthropicKey ? "on (" + cfg.model + ")" : "off: add ANTHROPIC_API_KEY to .env"}`);
    console.log(`  Job search       : ${cfg.adzunaId && cfg.adzunaKey ? "on (Adzuna)" : "off: add ADZUNA_APP_ID and ADZUNA_APP_KEY to .env"}`);
    console.log(`  Your data        : ${cfg.dataDir}${path.sep}tracker.json\n`);
    console.log("  Press Ctrl+C to stop.\n");
  });
}

module.exports = { createServer };
