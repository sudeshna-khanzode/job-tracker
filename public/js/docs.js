/* Document builders: tailored resume + cover letter as .docx (docx lib) and .pdf (jsPDF).
   Works in the browser (window.docx / window.jspdf) and in node for testing. */
(function (root) {
  "use strict";

  const clean = (s) => String(s == null ? "" : s).replace(/[ \t]+/g, " ").replace(/\s*\n\s*/g, " ").trim();
  const arr = (a) => (Array.isArray(a) ? a.filter((x) => x != null && x !== "") : a ? [a] : []);
  const isA4 = (country) => !["US", "CA"].includes(String(country || "").toUpperCase());

  function normalizeResume(r) {
    r = r || {};
    return {
      name: clean(r.name),
      headline: clean(r.headline),
      contact: arr(r.contact).map(clean).filter(Boolean),
      summary: clean(r.summary),
      skills: arr(r.skills).map((g) =>
        typeof g === "string" ? { group: "", items: [clean(g)] } : { group: clean(g.group), items: arr(g.items).map(clean).filter(Boolean) }
      ).filter((g) => g.items.length),
      experience: arr(r.experience).map((e) => ({
        title: clean(e.title), company: clean(e.company), location: clean(e.location), dates: clean(e.dates),
        bullets: arr(e.bullets).map(clean).filter(Boolean),
      })),
      education: arr(r.education).map((e) => ({
        qualification: clean(e.qualification), institution: clean(e.institution), dates: clean(e.dates), details: clean(e.details),
      })),
      certifications: arr(r.certifications).map(clean).filter(Boolean),
      projects: arr(r.projects).map((p) => ({ name: clean(p.name), description: clean(p.description), bullets: arr(p.bullets).map(clean).filter(Boolean) })),
      additional: arr(r.additional).map(clean).filter(Boolean),
    };
  }

  function normalizeLetter(c, r) {
    c = c || {};
    return {
      name: clean((r && r.name) || c.name),
      contact: arr((r && r.contact) || c.contact).map(clean).filter(Boolean),
      date: clean(c.date),
      recipient: arr(c.recipient).map(clean).filter(Boolean),
      subject: clean(c.subject),
      greeting: clean(c.greeting) || "Dear Hiring Manager,",
      paragraphs: arr(c.paragraphs).map((p) => String(p).trim()).filter(Boolean),
      closing: clean(c.closing) || "Kind regards,",
    };
  }

  /* ---------------- DOCX ---------------- */
  function resumeDocx(D, resume, opts) {
    const r = normalizeResume(resume);
    const a4 = isA4(opts && opts.country);
    const W = a4 ? 11906 : 12240, H = a4 ? 16838 : 15840, M = 850;
    const RIGHT = W - 2 * M;
    const ACCENT = "1F4E79";
    const kids = [];
    const P = (o) => kids.push(new D.Paragraph(o));
    const run = (text, o) => new D.TextRun(Object.assign({ text }, o || {}));

    P({ alignment: D.AlignmentType.CENTER, spacing: { after: 40 }, children: [run(r.name || "Your Name", { bold: true, size: 34, color: ACCENT })] });
    if (r.headline) P({ alignment: D.AlignmentType.CENTER, spacing: { after: 40 }, children: [run(r.headline, { size: 22, color: "444444" })] });
    if (r.contact.length) P({ alignment: D.AlignmentType.CENTER, spacing: { after: 120 }, children: [run(r.contact.join("  |  "), { size: 19, color: "444444" })] });

    const section = (title) => P({
      spacing: { before: 200, after: 80 }, keepNext: true, keepLines: true,
      border: { bottom: { color: "9AA5B1", space: 2, style: D.BorderStyle.SINGLE, size: 6 } },
      children: [run(title.toUpperCase(), { bold: true, size: 21, color: ACCENT, characterSpacing: 20 })],
    });
    const bullet = (text) => P({ bullet: { level: 0 }, spacing: { after: 30 }, children: [run(text, { size: 20 })] });
    const titleLine = (left, leftRest, right) => P({
      tabStops: [{ type: D.TabStopType.RIGHT, position: RIGHT }],
      spacing: { before: 120, after: 20 }, keepNext: true,
      children: [run(left, { bold: true, size: 21 }), run(leftRest ? "  ·  " + leftRest : "", { size: 21 }), run(right ? "\t" + right : "", { size: 20, color: "444444" })].filter(Boolean),
    });

    if (r.summary) { section("Professional Summary"); P({ spacing: { after: 60 }, children: [run(r.summary, { size: 20 })] }); }
    if (r.skills.length) {
      section("Key Skills");
      r.skills.forEach((g) => P({ spacing: { after: 40 }, children: [g.group ? run(g.group + ": ", { bold: true, size: 20 }) : run(""), run(g.items.join(", "), { size: 20 })] }));
    }
    if (r.experience.length) {
      section("Professional Experience");
      r.experience.forEach((e) => {
        titleLine(e.title, e.company, e.dates);
        if (e.location) P({ spacing: { after: 40 }, keepNext: true, children: [run(e.location, { italics: true, size: 19, color: "555555" })] });
        e.bullets.forEach(bullet);
      });
    }
    if (r.projects.length) {
      section("Projects");
      r.projects.forEach((p) => {
        P({ spacing: { before: 100, after: 20 }, keepNext: true, children: [run(p.name, { bold: true, size: 21 }), run(p.description ? " – " + p.description : "", { size: 20 })] });
        p.bullets.forEach(bullet);
      });
    }
    if (r.education.length) {
      section("Education");
      r.education.forEach((e) => {
        titleLine(e.qualification, e.institution, e.dates);
        if (e.details) P({ spacing: { after: 40 }, children: [run(e.details, { size: 20 })] });
      });
    }
    if (r.certifications.length) { section("Certifications & Licences"); r.certifications.forEach(bullet); }
    if (r.additional.length) { section("Additional Information"); r.additional.forEach(bullet); }

    return new D.Document({
      creator: r.name || "Job Tracker",
      title: (r.name || "Resume") + " - Resume",
      styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
      sections: [{ properties: { page: { size: { width: W, height: H }, margin: { top: 720, bottom: 720, left: M, right: M } } }, children: kids }],
    });
  }

  function letterDocx(D, letter, resume, opts) {
    const c = normalizeLetter(letter, normalizeResume(resume));
    const a4 = isA4(opts && opts.country);
    const W = a4 ? 11906 : 12240, H = a4 ? 16838 : 15840, M = 1134;
    const kids = [];
    const P = (o) => kids.push(new D.Paragraph(o));
    const run = (text, o) => new D.TextRun(Object.assign({ text, size: 22 }, o || {}));
    P({ spacing: { after: 20 }, children: [run(c.name || "Your Name", { bold: true, size: 30, color: "1F4E79" })] });
    if (c.contact.length) P({ spacing: { after: 240 }, border: { bottom: { color: "9AA5B1", space: 4, style: D.BorderStyle.SINGLE, size: 6 } }, children: [run(c.contact.join("  |  "), { size: 19, color: "444444" })] });
    if (c.date) P({ spacing: { after: 200 }, children: [run(c.date)] });
    c.recipient.forEach((l, i) => P({ spacing: { after: i === c.recipient.length - 1 ? 200 : 0 }, children: [run(l)] }));
    if (c.subject) P({ spacing: { after: 200 }, children: [run(c.subject, { bold: true })] });
    P({ spacing: { after: 160 }, children: [run(c.greeting)] });
    c.paragraphs.forEach((p) => P({ spacing: { after: 160, line: 276 }, children: [run(p)] }));
    P({ spacing: { before: 120, after: 480 }, children: [run(c.closing)] });
    P({ children: [run(c.name, { bold: true })] });
    return new D.Document({
      creator: c.name || "Job Tracker",
      title: (c.name || "") + " - Cover letter",
      styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
      sections: [{ properties: { page: { size: { width: W, height: H }, margin: { top: 1134, bottom: 1134, left: M, right: M } } }, children: kids }],
    });
  }

  /* ---------------- PDF ---------------- */
  function pdfSafe(s) {
    return String(s == null ? "" : s)
      .replace(/[‘’‚‛′]/g, "'").replace(/[“”„″]/g, '"')
      .replace(/[‐-―−]/g, "-").replace(/•|·/g, "-").replace(/…/g, "...")
      .replace(/[   ]/g, " ").replace(/[^\x20-\x7E¡-ÿ\n]/g, "");
  }

  function pdfWriter(jsPDF, country) {
    const doc = new jsPDF({ unit: "pt", format: isA4(country) ? "a4" : "letter" });
    const PW = doc.internal.pageSize.getWidth(), PH = doc.internal.pageSize.getHeight();
    const M = 50, BOTTOM = PH - 50, CW = PW - 2 * M;
    const st = { y: 56 };
    const ensure = (h) => { if (st.y + h > BOTTOM) { doc.addPage(); st.y = 56; } };
    const font = (size, bold, italic, color) => {
      doc.setFont("helvetica", bold && italic ? "bolditalic" : bold ? "bold" : italic ? "italic" : "normal");
      doc.setFontSize(size);
      const c = color || [20, 20, 20];
      doc.setTextColor(c[0], c[1], c[2]);
    };
    const para = (text, o) => {
      o = o || {};
      const size = o.size || 10, lh = size * (o.lh || 1.32), indent = o.indent || 0;
      font(size, o.bold, o.italic, o.color);
      const lines = doc.splitTextToSize(pdfSafe(text), CW - indent);
      lines.forEach((ln, i) => {
        ensure(lh);
        if (o.bullet && i === 0) { doc.setFillColor(60, 60, 60); doc.circle(M + indent - 7, st.y - size * 0.32, 1.4, "F"); }
        if (o.align === "center") doc.text(ln, PW / 2, st.y, { align: "center" });
        else doc.text(ln, M + indent, st.y);
        st.y += lh;
      });
      st.y += o.after == null ? 2 : o.after;
    };
    return { doc, st, M, CW, PW, ensure, font, para };
  }

  function resumePdf(jsPDF, resume, opts) {
    const r = normalizeResume(resume);
    const w = pdfWriter(jsPDF, opts && opts.country);
    const { doc, st, M, CW, PW } = w;
    const ACC = [31, 78, 121];
    w.para(r.name || "Your Name", { size: 19, bold: true, align: "center", color: ACC, after: 2 });
    if (r.headline) w.para(r.headline, { size: 11, align: "center", color: [70, 70, 70], after: 1 });
    if (r.contact.length) w.para(r.contact.join("  |  "), { size: 9, align: "center", color: [70, 70, 70], after: 6 });
    const section = (t) => {
      w.ensure(40); st.y += 8;
      w.font(10.5, true, false, ACC); doc.text(t.toUpperCase(), M, st.y);
      st.y += 4; doc.setDrawColor(154, 165, 177); doc.setLineWidth(0.6); doc.line(M, st.y, M + CW, st.y); st.y += 12;
    };
    const titleLine = (left, rest, right) => {
      w.ensure(34); st.y += 4;
      w.font(10, false, false, [70, 70, 70]);
      const rw = right ? doc.getTextWidth(pdfSafe(right)) : 0;
      if (right) doc.text(pdfSafe(right), M + CW, st.y, { align: "right" });
      w.font(10.5, true);
      const full = pdfSafe(left + (rest ? "  -  " + rest : ""));
      const lines = doc.splitTextToSize(full, CW - rw - 12);
      lines.forEach((ln) => { doc.text(ln, M, st.y); st.y += 13.5; });
    };
    const bullet = (t) => w.para(t, { size: 9.8, indent: 12, bullet: true, after: 1.5 });
    if (r.summary) { section("Professional Summary"); w.para(r.summary, { size: 9.8, after: 2 }); }
    if (r.skills.length) {
      section("Key Skills");
      r.skills.forEach((g) => w.para((g.group ? g.group + ": " : "") + g.items.join(", "), { size: 9.8, after: 2 }));
    }
    if (r.experience.length) {
      section("Professional Experience");
      r.experience.forEach((e) => {
        titleLine(e.title, e.company, e.dates);
        if (e.location) w.para(e.location, { size: 9, italic: true, color: [90, 90, 90], after: 2 });
        e.bullets.forEach(bullet);
      });
    }
    if (r.projects.length) {
      section("Projects");
      r.projects.forEach((p) => { titleLine(p.name, "", ""); if (p.description) w.para(p.description, { size: 9.8, after: 2 }); p.bullets.forEach(bullet); });
    }
    if (r.education.length) {
      section("Education");
      r.education.forEach((e) => { titleLine(e.qualification, e.institution, e.dates); if (e.details) w.para(e.details, { size: 9.8, after: 2 }); });
    }
    if (r.certifications.length) { section("Certifications & Licences"); r.certifications.forEach(bullet); }
    if (r.additional.length) { section("Additional Information"); r.additional.forEach(bullet); }
    void PW;
    return doc;
  }

  function letterPdf(jsPDF, letter, resume, opts) {
    const c = normalizeLetter(letter, normalizeResume(resume));
    const w = pdfWriter(jsPDF, opts && opts.country);
    const { doc, st, M, CW } = w;
    w.para(c.name || "Your Name", { size: 16, bold: true, color: [31, 78, 121], after: 2 });
    if (c.contact.length) {
      w.para(c.contact.join("  |  "), { size: 9, color: [70, 70, 70], after: 2 });
      doc.setDrawColor(154, 165, 177); doc.setLineWidth(0.6); doc.line(M, st.y, M + CW, st.y); st.y += 22;
    }
    if (c.date) w.para(c.date, { size: 10.5, after: 12 });
    c.recipient.forEach((l, i) => w.para(l, { size: 10.5, after: i === c.recipient.length - 1 ? 12 : 0 }));
    if (c.subject) w.para(c.subject, { size: 10.5, bold: true, after: 12 });
    w.para(c.greeting, { size: 10.5, after: 9 });
    c.paragraphs.forEach((p) => w.para(p, { size: 10.5, lh: 1.42, after: 9 }));
    st.y += 4; w.para(c.closing, { size: 10.5, after: 26 });
    w.para(c.name, { size: 10.5, bold: true });
    return doc;
  }

  /* ---------------- plain text (for copy) ---------------- */
  function resumeText(resume) {
    const r = normalizeResume(resume), out = [];
    out.push(r.name); if (r.headline) out.push(r.headline); if (r.contact.length) out.push(r.contact.join(" | "));
    const sec = (t) => out.push("", t.toUpperCase());
    if (r.summary) { sec("Professional Summary"); out.push(r.summary); }
    if (r.skills.length) { sec("Key Skills"); r.skills.forEach((g) => out.push((g.group ? g.group + ": " : "") + g.items.join(", "))); }
    if (r.experience.length) { sec("Professional Experience"); r.experience.forEach((e) => { out.push("", [e.title, e.company].filter(Boolean).join(" - ") + (e.dates ? " (" + e.dates + ")" : "")); if (e.location) out.push(e.location); e.bullets.forEach((b) => out.push("• " + b)); }); }
    if (r.projects.length) { sec("Projects"); r.projects.forEach((p) => { out.push("", p.name + (p.description ? " - " + p.description : "")); p.bullets.forEach((b) => out.push("• " + b)); }); }
    if (r.education.length) { sec("Education"); r.education.forEach((e) => { out.push([e.qualification, e.institution].filter(Boolean).join(" - ") + (e.dates ? " (" + e.dates + ")" : "")); if (e.details) out.push(e.details); }); }
    if (r.certifications.length) { sec("Certifications & Licences"); r.certifications.forEach((b) => out.push("• " + b)); }
    if (r.additional.length) { sec("Additional Information"); r.additional.forEach((b) => out.push("• " + b)); }
    return out.join("\n");
  }
  function letterText(letter, resume) {
    const c = normalizeLetter(letter, normalizeResume(resume));
    return [c.name, c.contact.join(" | "), "", c.date, "", ...c.recipient, "", c.subject, "", c.greeting, "", c.paragraphs.join("\n\n"), "", c.closing, c.name]
      .filter((x, i, a) => !(x === "" && a[i - 1] === "")).join("\n").trim();
  }

  const api = { normalizeResume, normalizeLetter, resumeDocx, letterDocx, resumePdf, letterPdf, resumeText, letterText, pdfSafe };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.JTDocs = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
