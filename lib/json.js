"use strict";
// Tolerant JSON extraction from a model reply: whole reply, a ``` fenced block,
// or the span from the first { / [ to the last } / ].
function parseJsonLoose(text) {
  const t = String(text || "").trim();
  const tries = [t];
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) tries.push(fence[1].trim());
  const firstObj = t.search(/[{[]/);
  const lastObj = Math.max(t.lastIndexOf("}"), t.lastIndexOf("]"));
  if (firstObj >= 0 && lastObj > firstObj) tries.push(t.slice(firstObj, lastObj + 1));
  for (const s of tries) {
    try { return { ok: true, value: JSON.parse(s) }; } catch {}
  }
  return { ok: false };
}
module.exports = { parseJsonLoose };
