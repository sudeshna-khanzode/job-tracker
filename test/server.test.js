"use strict";
// Runs the real server against fake Claude and Adzuna services.
const test = require("node:test");
const assert = require("node:assert");
const http = require("http");
const os = require("os");
const fs = require("fs");
const path = require("path");
const { createServer } = require("../server");
const { Store } = require("../lib/store");

const listen = (srv) => new Promise((r) => srv.listen(0, "127.0.0.1", () => r(srv.address().port)));

test("server endpoints", async (t) => {
  const seen = {};
  const fake = http.createServer((req, res) => {
    const u = new URL(req.url, "http://x");
    if (u.pathname === "/v1/messages") {
      let b = ""; req.on("data", (c) => (b += c)); req.on("end", () => {
        seen.claude = { headers: req.headers, body: JSON.parse(b) };
        if (req.headers["x-api-key"] !== "sk-test") { res.writeHead(401); return res.end(JSON.stringify({ error: { message: "bad key" } })); }
        res.writeHead(200, { "content-type": "application/json" });
        res.end(JSON.stringify({ model: "m", stop_reason: "end_turn", content: [{ type: "text", text: '{"ok":true}' }] }));
      });
      return;
    }
    if (u.pathname.startsWith("/v1/api/jobs/au/search/1")) {
      seen.adzuna = u.searchParams;
      res.writeHead(200, { "content-type": "application/json" });
      return res.end(JSON.stringify({ results: [{ id: 1, title: "Dev", company: { display_name: "Acme" }, location: { display_name: "Adelaide" }, redirect_url: "https://a/1", description: "desc", created: "2026-11-10T00:00:00Z" }] }));
    }
    res.writeHead(404); res.end();
  });
  const fport = await listen(fake);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "jt-"));
  const cfg = { anthropicKey: "sk-test", model: "claude-sonnet-5-5", fastModel: "claude-haiku-5-5", anthropicBase: `http://127.0.0.1:${fport}`,
    adzunaId: "id", adzunaKey: "key", adzunaBase: `http://127.0.0.1:${fport}`, dataDir: dir };
  const srv = createServer(cfg, new Store(dir));
  const port = await listen(srv);
  const base = `http://127.0.0.1:${port}`;
  t.after(() => { srv.close(); fake.close(); });

  const st = await (await fetch(base + "/api/status")).json();
  assert.strictEqual(st.claude, true); assert.strictEqual(st.jobs, true);

  const page = await fetch(base + "/");
  assert.strictEqual(page.status, 200);
  assert.match(await page.text(), /Job Tracker/);
  assert.strictEqual((await fetch(base + "/../server.js")).status, 404);

  const put = await fetch(base + "/api/store?path=profile/main", { method: "PUT", headers: { "content-type": "application/json" }, body: JSON.stringify({ skills: "SQL" }) });
  assert.strictEqual(put.status, 200);
  assert.deepStrictEqual((await (await fetch(base + "/api/store")).json())["profile/main"], { skills: "SQL" });
  assert.strictEqual((await fetch(base + "/api/store?path=../x", { method: "PUT", body: "{}" })).status, 400);

  const c = await (await fetch(base + "/api/claude", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ prompt: "hi", tier: "quick" }) })).json();
  assert.strictEqual(c.text, '{"ok":true}');
  assert.strictEqual(seen.claude.body.model, "claude-haiku-5-5");
  assert.strictEqual(seen.claude.headers["anthropic-version"], "2023-06-01");

  const j = await (await fetch(base + "/api/jobs/search?what=dev&where=Adelaide&country=au&jobType=fulltime")).json();
  assert.strictEqual(j.results[0].company, "Acme");
  assert.strictEqual(seen.adzuna.get("full_time"), "1");
  assert.strictEqual(seen.adzuna.get("where"), "Adelaide");

  const cross = await fetch(base + "/api/status", { headers: { origin: "https://evil.example" } });
  assert.strictEqual(cross.status, 403);
});
