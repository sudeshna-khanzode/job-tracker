"use strict";
// Job search through the Adzuna API (aggregates listings from many job boards).
// Docs: https://developer.adzuna.com/docs/search

const COUNTRIES = ["au", "nz", "gb", "us", "ca", "in", "sg", "za", "de", "nl", "fr", "it", "es", "at", "be", "ch", "pl", "br", "mx"];
const TYPE_PARAM = { fulltime: "full_time", parttime: "part_time", contract: "contract", permanent: "permanent" };

const fail = (code, message, status = 502) => Object.assign(new Error(message), { code, status });
const stripTags = (s) => String(s || "").replace(/<[^>]*>/g, "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#0?39;/g, "'").replace(/\s+/g, " ").trim();

function money(n) { return Math.round(n).toLocaleString("en"); }
const CUR = { gb: "£", de: "€", nl: "€", fr: "€", it: "€", es: "€", at: "€", be: "€", in: "₹", ch: "CHF ", pl: "PLN ", br: "R$", za: "R" };
function payText(r, cc) {
  const c = CUR[cc] || "$";
  if (!r.salary_min && !r.salary_max) return "";
  if (r.salary_is_predicted === "1" || r.salary_is_predicted === 1) return ""; // skip Adzuna's estimates
  if (r.salary_min && r.salary_max && Math.round(r.salary_min) !== Math.round(r.salary_max)) return `${c}${money(r.salary_min)} – ${c}${money(r.salary_max)}`;
  return `${c}${money(r.salary_min || r.salary_max)}`;
}
const TIME = { full_time: "Full-time", part_time: "Part-time" };
const CONTRACT = { contract: "Contract", permanent: "Permanent" };

function normalize(r, cc) {
  return {
    id: String(r.id || ""),
    title: stripTags(r.title),
    company: stripTags(r.company && r.company.display_name),
    location: stripTags(r.location && r.location.display_name),
    posted: r.created || "",
    jobType: [TIME[r.contract_time], CONTRACT[r.contract_type]].filter(Boolean).join(" · "),
    pay: payText(r, cc),
    url: /^https?:\/\//.test(r.redirect_url || "") ? r.redirect_url : "",
    snippet: stripTags(r.description).slice(0, 2000),
  };
}

async function searchJobs(cfg, { what, where, country = "au", jobType = "", maxDaysOld = 45, perPage = 20 }, signal) {
  if (!cfg.adzunaId || !cfg.adzunaKey) throw fail("no_keys", "ADZUNA_APP_ID and ADZUNA_APP_KEY are not set in .env", 400);
  const cc = String(country || "").toLowerCase();
  if (!COUNTRIES.includes(cc)) throw fail("bad_country", `Adzuna doesn't cover country "${country}". Supported: ${COUNTRIES.join(", ")}`, 400);
  if (!String(what || "").trim()) throw fail("bad_request", "what is required", 400);
  const q = new URLSearchParams({
    app_id: cfg.adzunaId, app_key: cfg.adzunaKey, what: String(what).trim(),
    results_per_page: String(Math.min(Math.max(Number(perPage) || 20, 1), 50)),
    max_days_old: String(maxDaysOld), sort_by: "relevance", "content-type": "application/json",
  });
  const wh = String(where || "").trim();
  if (wh && !/^remote$/i.test(wh)) q.set("where", wh);
  if (/^remote$/i.test(wh)) q.set("what", q.get("what") + " remote");
  if (TYPE_PARAM[jobType]) q.set(TYPE_PARAM[jobType], "1");
  let res;
  try {
    res = await fetch(`${cfg.adzunaBase}/v1/api/jobs/${cc}/search/1?${q}`, { signal, headers: { accept: "application/json" } });
  } catch (e) {
    if (e.name === "AbortError") throw fail("cancelled", "Cancelled", 499);
    throw fail("upstream_error", "Couldn't reach Adzuna: " + e.message);
  }
  const data = await res.json().catch(() => null);
  if (res.status === 401 || res.status === 403) throw fail("bad_keys", "Adzuna rejected your app ID/key. Check ADZUNA_APP_ID and ADZUNA_APP_KEY.", 401);
  if (res.status === 429) throw fail("rate_limited", "Adzuna rate limit reached. Wait a minute and try again.", 429);
  if (!res.ok || !data) throw fail("upstream_error", `Adzuna error ${res.status}`);
  return (data.results || []).map((r) => normalize(r, cc)).filter((r) => r.title);
}

module.exports = { searchJobs, normalize, COUNTRIES };
