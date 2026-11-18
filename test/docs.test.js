"use strict";
// Builds a resume and cover letter as .docx and .pdf with the bundled libraries.
const test = require("node:test");
const assert = require("node:assert");
const D = require("../public/vendor/docx.umd.js");
const { jsPDF } = require("../public/vendor/jspdf.umd.min.js");
const J = require("../public/js/docs.js");

const resume = { name: "Alex Nguyen", headline: "Data Analyst", contact: ["alex@example.com", "Adelaide SA"], summary: "Analyst with 4 years’ experience — reporting.",
  skills: [{ group: "Technical", items: ["SQL", "Power BI"] }], experience: [{ title: "Analyst", company: "Acme", dates: "2022 – Present", bullets: ["Built dashboards."] }],
  education: [{ qualification: "BSc", institution: "Uni", dates: "2019" }] };
const letter = { date: "18 November 2026", recipient: ["Hiring Manager", "Acme"], greeting: "Dear Hiring Manager,", paragraphs: ["One.", "Two."] };

test("docx files are valid zip archives", async () => {
  for (const d of [J.resumeDocx(D, resume, { country: "AU" }), J.letterDocx(D, letter, resume, { country: "US" })]) {
    const buf = await D.Packer.toBuffer(d);
    assert.strictEqual(buf.slice(0, 2).toString(), "PK");
    assert.ok(buf.length > 5000);
  }
});

test("pdf files start with a PDF header", () => {
  for (const p of [J.resumePdf(jsPDF, resume, { country: "AU" }), J.letterPdf(jsPDF, letter, resume, {})]) {
    const out = Buffer.from(p.output("arraybuffer"));
    assert.strictEqual(out.slice(0, 5).toString(), "%PDF-");
  }
});

test("plain text versions include the key content", () => {
  assert.match(J.resumeText(resume), /ALEX|Alex Nguyen/);
  assert.match(J.resumeText(resume), /• Built dashboards\./);
  assert.match(J.letterText(letter, resume), /Dear Hiring Manager,[\s\S]*Two\./);
});
