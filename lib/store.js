"use strict";
// Tiny document store: one JSON file on disk, keyed by slash-separated paths
// (e.g. "profile/main", "profile/main/jobs/j-abc"). Writes are atomic.
const fs = require("fs");
const path = require("path");

const PATH_RE = /^[A-Za-z0-9_\-.~:@+]+(\/[A-Za-z0-9_\-.~:@+]+)*$/;
const validPath = (p) => typeof p === "string" && p.length <= 400 && PATH_RE.test(p) && !p.split("/").some((s) => s === "." || s === "..");

class Store {
  constructor(dir) {
    this.dir = dir;
    this.file = path.join(dir, "tracker.json");
    this.docs = {};
    this.writing = Promise.resolve();
    try {
      this.docs = JSON.parse(fs.readFileSync(this.file, "utf8")) || {};
    } catch (e) {
      if (e.code !== "ENOENT") {
        // keep a copy of an unreadable file rather than overwrite it
        const bad = this.file + ".unreadable-" + Date.now();
        try { fs.copyFileSync(this.file, bad); } catch {}
        console.warn(`Couldn't read ${this.file} (${e.message}). Saved a copy to ${bad} and started fresh.`);
      }
      this.docs = {};
    }
  }
  all() { return this.docs; }
  get(p) { return this.docs[p]; }
  set(p, data) {
    if (!validPath(p)) throw Object.assign(new Error("Invalid path"), { status: 400 });
    if (!data || typeof data !== "object" || Array.isArray(data)) throw Object.assign(new Error("Body must be a JSON object"), { status: 400 });
    this.docs[p] = data;
    return this.flush();
  }
  delete(p) {
    if (!validPath(p)) throw Object.assign(new Error("Invalid path"), { status: 400 });
    delete this.docs[p];
    return this.flush();
  }
  flush() {
    const snapshot = JSON.stringify(this.docs, null, 1);
    this.writing = this.writing.catch(() => {}).then(async () => {
      await fs.promises.mkdir(this.dir, { recursive: true });
      const tmp = this.file + ".tmp";
      await fs.promises.writeFile(tmp, snapshot);
      await fs.promises.rename(tmp, this.file);
    });
    return this.writing;
  }
}

module.exports = { Store, validPath };
