"use strict";
const test = require("node:test");
const assert = require("node:assert");
const os = require("os");
const fs = require("fs");
const path = require("path");
const { parseJsonLoose } = require("../lib/json");
const { normalize } = require("../lib/adzuna");
const { Store, validPath } = require("../lib/store");

test("parseJsonLoose handles plain, fenced and wrapped JSON", () => {
  assert.deepStrictEqual(parseJsonLoose('{"a":1}').value, { a: 1 });
  assert.deepStrictEqual(parseJsonLoose('Here:\n```json\n[1,2]\n```').value, [1, 2]);
  assert.deepStrictEqual(parseJsonLoose('Sure! {"roles":["A"]} Hope that helps.').value, { roles: ["A"] });
  assert.strictEqual(parseJsonLoose("no json here").ok, false);
});

test("adzuna results are normalised and cleaned", () => {
  const r = normalize({ id: 7, title: "<strong>Data</strong> Analyst", company: { display_name: "A &amp; B" }, location: { display_name: "Adelaide, SA" },
    created: "2026-11-10T00:00:00Z", contract_time: "full_time", contract_type: "permanent", salary_min: 90000, salary_max: 100000,
    redirect_url: "https://www.adzuna.com.au/land/ad/7", description: "Great <b>role</b>" }, "au");
  assert.strictEqual(r.title, "Data Analyst");
  assert.strictEqual(r.company, "A & B");
  assert.strictEqual(r.jobType, "Full-time · Permanent");
  assert.strictEqual(r.pay, "$90,000 – $100,000");
  assert.strictEqual(r.snippet, "Great role");
  assert.strictEqual(normalize({ title: "x", redirect_url: "javascript:alert(1)" }, "au").url, "");
});

test("store validates paths and persists to disk", async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "jt-"));
  const s = new Store(dir);
  assert.ok(validPath("profile/main/jobs/j-1"));
  assert.ok(!validPath("../etc/passwd"));
  assert.ok(!validPath("a//b"));
  await s.set("profile/main", { skills: "SQL" });
  assert.deepStrictEqual(new Store(dir).get("profile/main"), { skills: "SQL" });
  await s.delete("profile/main");
  assert.strictEqual(new Store(dir).get("profile/main"), undefined);
  assert.throws(() => s.set("bad path", {}), /Invalid path/);
});
