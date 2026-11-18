/* Local platform layer: gives app.js its storage, Claude, job search and downloads
   by talking to the local server (server.js). */
(() => {
  "use strict";

  const api = async (url, opts = {}) => {
    let res;
    try { res = await fetch(url, opts); }
    catch (e) {
      if (e.name === "AbortError") throw { code: "cancelled", message: "Cancelled" };
      throw { code: "offline", message: "Can't reach the Job Tracker server. Is it still running?" };
    }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw Object.assign({ code: "server_error", message: "Server error " + res.status }, data.error || {});
    return data;
  };

  /* ---------- db: in-memory copy of ./data/tracker.json, saved through /api/store ---------- */
  const docs = new Map();
  const listeners = new Set();
  let loadErr = null;
  const loaded = api("/api/store").then((all) => { Object.entries(all || {}).forEach(([k, v]) => docs.set(k, v)); }, (e) => { loadErr = e; });
  const meta = { fromCache: false, hasPendingWrites: false };
  const notify = () => queueMicrotask(() => listeners.forEach((f) => { try { f(); } catch (e) { console.error(e); } }));
  const snap = (p) => ({ id: p.split("/").pop(), exists: docs.has(p), data: () => docs.get(p), metadata: meta });
  const children = (col) => {
    const depth = col.split("/").length + 1, pre = col + "/";
    return [...docs.keys()].filter((k) => k.startsWith(pre) && k.split("/").length === depth).sort().map(snap);
  };
  const listen = (fn, err) => {
    loaded.then(() => { if (loadErr) { err && err({ code: "unavailable", message: loadErr.message }); return; } fn(); listeners.add(fn); });
    return () => listeners.delete(fn);
  };
  const enc = encodeURIComponent;
  function doc(p) {
    return {
      id: p.split("/").pop(), path: p,
      get: async () => { await loaded; return snap(p); },
      set: async (data) => {
        const copy = JSON.parse(JSON.stringify(data));
        docs.set(p, copy); notify();
        await api("/api/store?path=" + enc(p), { method: "PUT", headers: { "content-type": "application/json" }, body: JSON.stringify(copy) });
      },
      update: async (data) => { await loaded; if (!docs.has(p)) throw { code: "invalid_argument", message: "No such document" }; return doc(p).set(Object.assign({}, docs.get(p), data)); },
      delete: async () => { docs.delete(p); notify(); await api("/api/store?path=" + enc(p), { method: "DELETE" }); },
      onSnapshot: (next, err) => listen(() => next(snap(p)), err),
      collection: (c) => collection(p + "/" + c),
    };
  }
  function collection(p) {
    return {
      path: p,
      doc: (id) => doc(p + "/" + id),
      onSnapshot: (next, err) => listen(() => { const d = children(p); next({ docs: d, size: d.length, empty: !d.length, metadata: meta }); }, err),
    };
  }
  const db = { doc, collection };

  /* ---------- user ---------- */
  const user = { id: async () => "local", isOwner: () => true };

  /* ---------- sample: ask Claude through /api/claude ---------- */
  function parseJsonLoose(text) {
    const t = String(text || "").trim(), tries = [t];
    const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i); if (fence) tries.push(fence[1].trim());
    const a = t.search(/[{[]/), b = Math.max(t.lastIndexOf("}"), t.lastIndexOf("]"));
    if (a >= 0 && b > a) tries.push(t.slice(a, b + 1));
    for (const s of tries) { try { return { ok: true, value: JSON.parse(s) }; } catch {} }
    return { ok: false };
  }
  async function ask(input, opts = {}, json = false) {
    const prompt = Array.isArray(input) ? input.map((m) => m.content).join("\n\n") : String(input);
    const body = { prompt: json ? prompt + "\n\nRespond with JSON only, no other text." : prompt, tier: opts.modelTier || "default" };
    const out = await api("/api/claude", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body), signal: opts.signal });
    if (typeof opts.onText === "function") { try { opts.onText({ text: out.text, delta: out.text }); } catch {} }
    return out;
  }
  const sample = async (input, opts) => { const o = await ask(input, opts); return { text: o.text, truncated: !!o.truncated }; };
  sample.json = async (input, opts) => {
    const o = await ask(input, opts, true);
    const p = parseJsonLoose(o.text);
    if (!p.ok || o.truncated) throw { code: "invalid_json", message: o.truncated ? "Answer was cut short" : "No JSON in answer", text: o.text };
    return p.value;
  };

  /* ---------- downloads ---------- */
  const downloads = {
    save: async ({ filename, data }) => {
      const blob = data instanceof Blob ? data : new Blob([data], { type: filename.endsWith(".csv") ? "text/csv" : "text/plain" });
      const url = URL.createObjectURL(blob);
      const a = Object.assign(document.createElement("a"), { href: url, download: filename });
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      return { status: "saved" };
    },
  };

  /* ---------- job search (Adzuna via /api/jobs/search) ---------- */
  const jobs = {
    search: async (q, signal) => (await api("/api/jobs/search?" + new URLSearchParams(q), { signal })).results || [],
  };

  /* ---------- server status ---------- */
  const status = api("/api/status").catch(() => null);

  const caps = { db, user, sample, downloads, jobs, status: { get: () => status } };
  window.claude = { use: async (name) => caps[name] || null };
})();
