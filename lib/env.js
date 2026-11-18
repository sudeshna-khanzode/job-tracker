"use strict";
// Minimal .env loader (no dependencies). Existing environment variables win.
const fs = require("fs");
const path = require("path");

function loadEnv(file = path.join(__dirname, "..", ".env")) {
  let text;
  try { text = fs.readFileSync(file, "utf8"); } catch { return {}; }
  const out = {};
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    const m = line.match(/^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!m) continue;
    let v = m[2].trim();
    if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
    else v = v.replace(/\s+#.*$/, "");
    out[m[1]] = v;
    if (process.env[m[1]] === undefined) process.env[m[1]] = v;
  }
  return out;
}

const config = () => ({
  port: Number(process.env.PORT) || 3000,
  host: process.env.HOST || "127.0.0.1",
  anthropicKey: (process.env.ANTHROPIC_API_KEY || "").trim(),
  model: (process.env.ANTHROPIC_MODEL || "claude-sonnet-5-5").trim(),
  fastModel: (process.env.ANTHROPIC_FAST_MODEL || "claude-haiku-5-5").trim(),
  anthropicBase: (process.env.ANTHROPIC_BASE_URL || "https://api.anthropic.com").replace(/\/+$/, ""),
  adzunaId: (process.env.ADZUNA_APP_ID || "").trim(),
  adzunaKey: (process.env.ADZUNA_APP_KEY || "").trim(),
  adzunaBase: (process.env.ADZUNA_BASE_URL || "https://api.adzuna.com").replace(/\/+$/, ""),
  dataDir: process.env.DATA_DIR || path.join(__dirname, "..", "data"),
});

module.exports = { loadEnv, config };
