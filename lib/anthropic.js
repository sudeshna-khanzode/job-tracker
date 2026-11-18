"use strict";
// Calls the Claude Messages API. Your API key stays on this server; the browser never sees it.

const fail = (code, message, status = 502) => Object.assign(new Error(message), { code, status });

async function askClaude(cfg, { prompt, tier = "default", maxTokens }, signal) {
  if (!cfg.anthropicKey) throw fail("no_api_key", "ANTHROPIC_API_KEY is not set in .env", 400);
  if (typeof prompt !== "string" || !prompt.trim()) throw fail("bad_request", "prompt is required", 400);
  if (prompt.length > 400000) throw fail("prompt_too_large", "Prompt is too long", 413);
  const quick = tier === "quick";
  const body = {
    model: quick ? cfg.fastModel : cfg.model,
    max_tokens: Math.min(Number(maxTokens) || (quick ? 2000 : 8000), 16000),
    messages: [{ role: "user", content: prompt }],
  };
  let res;
  try {
    res = await fetch(cfg.anthropicBase + "/v1/messages", {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": cfg.anthropicKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify(body),
      signal,
    });
  } catch (e) {
    if (e.name === "AbortError") throw fail("cancelled", "Cancelled", 499);
    throw fail("upstream_error", "Couldn't reach the Claude API: " + e.message);
  }
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = (data && data.error && data.error.message) || res.statusText;
    if (res.status === 401 || res.status === 403) throw fail("invalid_key", "The Claude API rejected your API key: " + msg, 401);
    if (res.status === 429) throw fail("rate_limited", "Claude API rate limit: " + msg, 429);
    if (res.status === 413) throw fail("prompt_too_large", msg, 413);
    if (res.status === 400) throw fail("bad_request", msg, 400);
    if (res.status === 404) throw fail("bad_model", `Model not found (${body.model}). Check ANTHROPIC_MODEL in .env. ${msg}`, 400);
    throw fail("upstream_error", `Claude API error ${res.status}: ${msg}`);
  }
  const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
  if (!text.trim()) throw fail("empty_completion", "Claude returned no text");
  return { text, truncated: data.stop_reason === "max_tokens", model: data.model, usage: data.usage };
}

module.exports = { askClaude };
