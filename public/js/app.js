/* Job Application Tracker — app logic */
(() => {
"use strict";

/* ============ constants ============ */
const STATUSES = [
  ["shortlisted", "To apply"],
  ["ready", "Docs ready"],
  ["applied", "Applied"],
  ["response", "Got a response"],
  ["interview", "Interview"],
  ["offer", "Offer"],
  ["rejected", "Rejected"],
  ["withdrawn", "Withdrawn"],
  ["dismissed", "Not interested"],
];
const SLABEL = Object.fromEntries(STATUSES);
const TO_APPLY = new Set(["shortlisted", "ready"]);
const APPLIED_SET = new Set(["applied", "response", "interview", "offer", "rejected", "withdrawn"]);
const HEARD_SET = new Set(["response", "interview", "offer", "rejected"]);
const STATUS_PILL = { shortlisted: "", ready: "acc", applied: "acc", response: "warn", interview: "good", offer: "good", rejected: "bad", withdrawn: "", dismissed: "" };
const COUNTRIES = [["AU","Australia"],["NZ","New Zealand"],["US","United States"],["GB","United Kingdom"],["CA","Canada"],["IN","India"],["SG","Singapore"],["ZA","South Africa"],["DE","Germany"],["NL","Netherlands"],["FR","France"],["IT","Italy"],["ES","Spain"],["AT","Austria"],["BE","Belgium"],["CH","Switzerland"],["PL","Poland"],["BR","Brazil"],["MX","Mexico"]];
const JOBTYPES = [["","Any type"],["fulltime","Full-time"],["parttime","Part-time"],["contract","Contract"],["permanent","Permanent"]];
const ROUTES = ["home", "profile", "find", "apply", "tracker"];

/* ============ state ============ */
const S = {
  route: null,
  db: null, user: null, uid: null, sample: null, board: null, downloads: null, status: null,
  persist: "loading", // loading | ok | none | slow
  profile: { skills: "", notes: "", prefs: { roles: "", location: "", country: "AU", jobType: "" } },
  profileDirty: false,
  profileMsg: null, // {text, kind} shown beside the Save button
  resumes: new Map(),
  jobs: new Map(),
  loaded: { profile: false, jobs: false, resumes: false },
  find: { busy: false, msg: "", err: "", results: null, hidden: 0, selected: new Set(), count: 15, adding: false },
  tracker: { q: "", filter: "all", sort: "added" },
  ws: null,
  pendingAsk: null,
  renderDirty: false,
};

/* ============ utils ============ */
const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const safeUrl = (u) => (/^https?:\/\//i.test(String(u || "").trim()) ? String(u).trim() : "");
const norm = (s) => String(s || "").toLowerCase().replace(/&/g, " and ").replace(/[^a-z0-9]+/g, " ").trim();
const normCo = (s) => norm(s).replace(/\b(pty|ltd|limited|inc|llc|plc|group|the|corp|corporation|co)\b/g, " ").replace(/\s+/g, " ").trim();
const hash = (s) => { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0; return h.toString(36); };
const keyFor = (title, company) => "j-" + hash(norm(title) + "|" + normCo(company));
const now = () => Date.now();
const fmtDate = (t) => (t ? new Date(t).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" }) : "");
const isoDate = (t) => { if (!t) return ""; const d = new Date(t); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };
const daysSince = (t) => (t ? Math.floor((now() - t) / 86400000) : null);
const ago = (t) => { const d = daysSince(t); if (d == null) return ""; if (d <= 0) return "today"; if (d === 1) return "yesterday"; if (d < 14) return d + " days ago"; if (d < 60) return Math.round(d / 7) + " weeks ago"; return Math.round(d / 30) + " months ago"; };
const clone = (o) => JSON.parse(JSON.stringify(o));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const plural = (n, w) => `${n} ${w}${n === 1 ? "" : "s"}`;

function toast(msg, ms = 3200) {
  const t = $("#toast"); t.textContent = msg; t.hidden = false;
  clearTimeout(toast._t); toast._t = setTimeout(() => (t.hidden = true), ms);
}

const scripts = {};
function loadScript(src) {
  if (!scripts[src]) scripts[src] = new Promise((res, rej) => {
    const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => { delete scripts[src]; rej(new Error("Couldn't load " + src)); };
    document.head.appendChild(s);
  });
  return scripts[src];
}
const loadDocLibs = () => Promise.all([loadScript("vendor/docx.umd.js"), loadScript("vendor/jspdf.umd.min.js")]);

/* ============ persistence ============ */
const chains = new Map();
function queueWrite(id, fn) {
  const next = (chains.get(id) || Promise.resolve()).catch(() => {}).then(fn);
  chains.set(id, next);
  setSaveState("saving");
  next.then(() => setSaveState("saved"), (e) => { setSaveState("error"); dbError(e); });
  return next;
}
function dbError(e) {
  const c = e && e.code;
  if (c === "quota_exceeded") toast("Storage is full. Delete some old jobs or resumes and try again.", 6000);
  else if (c === "invalid_argument") toast("Couldn't save that change (it may be too large).", 6000);
  else if (c === "revoked") toast("Saving was turned off for this view.", 6000);
  else toast("Couldn't save just now. Your change is shown here but may not be stored.", 6000);
}
let saveStateT;
function setSaveState(s) {
  const el = $("#savestate");
  if (S.persist === "none") { el.textContent = "Not saving in this view"; return; }
  clearTimeout(saveStateT);
  if (s === "saving") el.textContent = "Saving…";
  else if (s === "saved") { el.textContent = "All changes saved"; }
  else if (s === "error") el.textContent = "Last save failed";
}
const profileRef = () => S.db.doc(`data/users/${S.uid}/profile`);
const jobsCol = () => profileRef().collection("jobs");
const resumesCol = () => profileRef().collection("resumes");

function saveProfile() {
  S.profileDirty = false;
  if (!S.db) return Promise.resolve();
  const body = clone(S.profile);
  return queueWrite("profile", () => profileRef().set(body));
}
function saveJob(job) {
  job.updatedAt = now();
  S.jobs.set(job.key, job);
  scheduleRender();
  if (!S.db) return Promise.resolve();
  const body = clone(job);
  return queueWrite("job:" + job.key, () => jobsCol().doc(job.key).set(body));
}
function deleteJob(key) {
  S.jobs.delete(key); scheduleRender();
  if (!S.db) return;
  queueWrite("job:" + key, () => jobsCol().doc(key).delete());
}
function saveResume(r) {
  S.resumes.set(r.id, r); scheduleRender();
  if (!S.db) return Promise.resolve();
  const body = clone(r);
  return queueWrite("resume:" + r.id, () => resumesCol().doc(r.id).set(body));
}
function deleteResume(id) {
  S.resumes.delete(id); scheduleRender();
  if (!S.db) return;
  queueWrite("resume:" + id, () => resumesCol().doc(id).delete());
}

function setStatus(key, status) {
  const j = S.jobs.get(key); if (!j || j.status === status) return;
  const job = clone(j);
  job.status = status;
  job.history = (job.history || []).concat([{ s: status, at: now() }]).slice(-30);
  if (APPLIED_SET.has(status) && !job.appliedAt) job.appliedAt = now();
  if (HEARD_SET.has(status) && !job.respondedAt) job.respondedAt = now();
  if (APPLIED_SET.has(status) || status === "dismissed") job.openedAt = null;
  saveJob(job);
}

/* ============ boot ============ */
const use = (n) => (window.claude && window.claude.use ? window.claude.use(n).catch(() => null) : Promise.resolve(null));

async function boot() {
  S.route = routeFromHash();
  renderChrome();
  const [db, user, sample, board, downloads, status] = await Promise.all(["db", "user", "sample", "jobs", "downloads", "status"].map(use));
  Object.assign(S, { db, user, sample, board, downloads });
  S.status = status ? await status.get() : null;
  S.uid = user ? await user.id().catch(() => null) : null;
  if (!db || !S.uid) {
    S.db = null; S.persist = "none";
    Object.keys(S.loaded).forEach((k) => (S.loaded[k] = true));
    setSaveState();
    afterLoad();
    return;
  }
  const timer = setTimeout(() => { if (S.persist === "loading") { S.persist = "slow"; Object.keys(S.loaded).forEach((k) => (S.loaded[k] = true)); afterLoad(); } }, 10000);
  const done = (k) => { S.loaded[k] = true; if (Object.values(S.loaded).every(Boolean)) { clearTimeout(timer); if (S.persist === "loading" || S.persist === "slow") { S.persist = "ok"; afterLoad(); } } };
  const fail = (k) => (e) => { console.warn("db", k, e); S.loaded[k] = true; if (e && e.code === "revoked") { S.persist = "none"; } done(k); scheduleRender(); };
  try {
    profileRef().onSnapshot((snap) => {
      if (snap.exists && !S.profileDirty) {
        const d = snap.data() || {};
        S.profile = { skills: d.skills || "", notes: d.notes || "", prefs: Object.assign({ roles: "", location: "", country: "AU", jobType: "" }, d.prefs || {}) };
      }
      if (!snap.metadata.fromCache || snap.exists) done("profile");
      scheduleRender();
    }, fail("profile"));
    jobsCol().onSnapshot((qs) => {
      const m = new Map(); qs.docs.forEach((d) => { const v = d.data(); if (v && v.key) m.set(d.id, v); });
      S.jobs = m;
      if (!qs.metadata.fromCache || qs.size) done("jobs");
      scheduleRender();
    }, fail("jobs"));
    resumesCol().onSnapshot((qs) => {
      const m = new Map(); qs.docs.forEach((d) => { const v = d.data(); if (v && v.id) m.set(d.id, v); });
      S.resumes = m;
      if (!qs.metadata.fromCache || qs.size) done("resumes");
      scheduleRender();
    }, fail("resumes"));
  } catch (e) {
    console.warn(e); S.db = null; S.persist = "none"; afterLoad();
  }
}
function afterLoad() {
  if (!location.hash) { S.route = "home"; history.replaceState(null, "", "#home"); }
  setSaveState(S.persist === "ok" ? "saved" : undefined);
  renderAll();
}

/* ============ routing & render ============ */
function routeFromHash() { const h = location.hash.replace("#", ""); return ROUTES.includes(h) ? h : "home"; }
window.addEventListener("hashchange", () => { S.route = routeFromHash(); renderAll(); window.scrollTo(0, 0); });

function isEditing(container) {
  const a = document.activeElement;
  if (!a || !container.contains(a)) return false;
  return a.tagName === "TEXTAREA" || a.tagName === "SELECT" || (a.tagName === "INPUT" && !["checkbox", "button", "file", "radio"].includes(a.type));
}
let rafPending = false;
function scheduleRender() {
  if (rafPending) return; rafPending = true;
  requestAnimationFrame(() => {
    rafPending = false;
    renderChrome();
    if (isEditing($("#view"))) { S.renderDirty = true; return; }
    renderView();
  });
}
let pointerDown = false;
const flushRender = () => { if (S.renderDirty && !pointerDown && !isEditing($("#view"))) { S.renderDirty = false; renderView(); } };
document.addEventListener("pointerdown", () => { pointerDown = true; }, true);
document.addEventListener("pointerup", () => setTimeout(() => { pointerDown = false; flushRender(); }, 60), true);
document.addEventListener("focusout", () => setTimeout(flushRender, 0));
function renderAll() { renderChrome(); renderView(); }

function counts() {
  let toApply = 0, ask = 0;
  for (const j of S.jobs.values()) { if (TO_APPLY.has(j.status)) { toApply++; if (j.openedAt) ask++; } }
  return { toApply, ask };
}
function renderChrome() {
  const c = counts();
  const tabs = [["home", "Home"], ["profile", "My profile"], ["find", "Find jobs"], ["apply", "Apply", c.toApply], ["tracker", "Tracker"]];
  $("#tabs").innerHTML = tabs.map(([r, l, n]) => `<a href="#${r}" class="${S.route === r ? "on" : ""}">${l}${n ? `<span class="n">${n}</span>` : ""}</a>`).join("");
  let b = "";
  if (S.persist === "none") b = `<div class="notice warn" style="margin-top:16px"><b>Can't reach the Job Tracker server.</b> Make sure it's running (<code>npm start</code>) and reload this page. Changes made now won't be saved.</div>`;
  else if (S.persist === "slow") b = `<div class="notice warn" style="margin-top:16px">Your saved data hasn't loaded yet. What you see may be out of date.</div>`;
  if ($("#banner").innerHTML !== b) $("#banner").innerHTML = b;
}
function renderView() {
  const v = $("#view");
  if (S.persist === "loading") { v.innerHTML = `<div class="progress"><div class="spin"></div>Loading your tracker…</div>`; return; }
  const scroll = window.scrollY;
  v.innerHTML = S.route === "home" ? viewHome() : S.route === "profile" ? viewProfile() : S.route === "apply" ? viewApply() : S.route === "tracker" ? viewTracker() : viewFind();
  window.scrollTo(0, scroll);
}


/* ============ HOME ============ */
function setupCard() {
  const st = S.status;
  if (!st) return `<div class="notice bad"><b>Can't reach the Job Tracker server.</b> Start it with <code>npm start</code> in the project folder, then reload this page.</div>`;
  if (st.claude && st.jobs) return "";
  const item = (ok, title, how) => `<li class="${ok ? "ok" : ""}"><span class="tick">${ok ? "✓" : "!"}</span><div><b>${title}</b><div class="small muted">${ok ? "Connected" : how}</div></div></li>`;
  return `<div class="card setup"><h2>Finish setting up</h2><p class="hint">Add these keys to the <code>.env</code> file in the project folder, then restart the server. The README has step-by-step instructions.</p>
    <ul class="setuplist">
      ${item(st.claude, "Claude API key: tailors your resume and cover letter", "Get a key at console.anthropic.com and set <code>ANTHROPIC_API_KEY</code>")}
      ${item(st.jobs, "Adzuna keys: finds jobs (free)", "Register at developer.adzuna.com and set <code>ADZUNA_APP_ID</code> and <code>ADZUNA_APP_KEY</code>")}
    </ul></div>`;
}
function viewHome() {
  const pr = S.profile.prefs, st = stats();
  const profileDone = S.resumes.size > 0 && !!String(pr.roles || "").trim() && !!String(pr.location || "").trim();
  const foundDone = st.live.length > 0;
  const appliedDone = st.applied.length > 0;
  const heardDone = st.heard.length > 0;
  const done = [profileDone, foundDone, appliedDone, heardDone];
  const next = done.indexOf(false);
  const followUps = st.waiting.filter((j) => daysSince(j.appliedAt) >= 14).length;
  const ready = st.toApply.filter((j) => j.status === "ready").length;
  const cta = next === 0 ? ["#profile", "Start with your profile"] : next === 1 ? ["#find", "Find your first jobs"] : st.toApply.length ? ["#apply", "Continue applying"] : ["#find", "Find more jobs"];
  const profileMissing = [!S.resumes.size && "a resume", !String(pr.roles || "").trim() && "job titles", !String(pr.location || "").trim() && "a location"].filter(Boolean);
  const steps = [
    { n: 1, href: "#profile", btn: "Go to My profile", title: "Set up your profile",
      text: "Upload your resume (PDF or Word), add your skills, and tell it which job titles and location to search.",
      detail: profileDone ? `${plural(S.resumes.size, "resume")} · searching “${esc(String(pr.roles).split(",")[0].trim())}”${String(pr.roles).includes(",") ? " and more" : ""}` : `Still needed: ${profileMissing.join(", ")}` },
    { n: 2, href: "#find", btn: "Find jobs", title: "Find jobs that fit",
      text: "Choose how many (10, 15 or 20). Jobs you've already tracked are skipped, and the rest are ranked against your resume. Tick the ones you like.",
      detail: foundDone ? `${plural(st.live.length, "job")} picked so far` : "Nothing picked yet" },
    { n: 3, href: "#apply", btn: "Open Apply", title: "Tailor &amp; apply",
      text: "Get a resume and cover letter written for each job, download them as Word or PDF, then apply on the job site. When you come back, it asks “Did you apply?”",
      detail: st.toApply.length ? `${plural(st.toApply.length, "job")} waiting${ready ? ` · ${ready} with documents ready` : ""}` : appliedDone ? "All caught up" : "No jobs waiting yet" },
    { n: 4, href: "#tracker", btn: "Open Tracker", title: "Track every application",
      text: "See everything you've applied to in one place. When a company replies, change the status: response, interview, offer or rejected.",
      detail: appliedDone ? `${plural(st.applied.length, "application")} · ${plural(st.heard.length, "reply", )}`.replace("replys", "replies") : "Nothing applied yet" },
  ];
  const state = (i) => (done[i] ? "done" : i === next ? "next" : "later");
  const name = (() => { const r = [...S.resumes.values()][0]; const m = r && String(r.text || "").trim().match(/^([A-Z][a-zA-Z'’-]+)\s+[A-Z]/); return m ? m[1] : ""; })();
  return `
  <section class="hero">
    <div class="hero-text">
      <div class="eyebrow">${name ? `Welcome back, ${esc(name)}` : "Welcome"}</div>
      <h1 class="hero-h">Your job hunt, in one calm place.</h1>
      <p class="hero-p">Find jobs that suit you, send a tailored resume and cover letter for each one, and always know where every application stands.</p>
      <div class="hero-cta"><a class="btn pri lg" href="${cta[0]}">${cta[1]} →</a>${next !== 0 ? `<a class="btn lg soft" href="#tracker">View tracker</a>` : ""}</div>
    </div>
    <div class="hero-stats">
      ${[["To apply", st.toApply.length, "#apply"], ["Applied", st.applied.length, "#tracker"], ["Waiting to hear", st.waiting.length, "#tracker"], ["Interviews", st.interviews.length, "#tracker"]]
        .map(([l, v, h]) => `<a class="hs" href="${h}"><span class="v">${v}</span><span class="l">${l}</span></a>`).join("")}
    </div>
  </section>
  ${setupCard()}
  ${followUps ? `<div class="notice warn" style="display:flex;align-items:center;gap:10px"><div style="flex:1"><b>${plural(followUps, "application")}</b> with no reply after 2 weeks. A short, polite follow-up email can help.</div><a class="btn sm" href="#tracker">See them</a></div>` : ""}
  <div class="howhead"><h2>How it works</h2><span class="muted small">${done.filter(Boolean).length} of 4 steps started</span></div>
  <div class="steps">
    ${steps.map((s, i) => `
      <div class="stepcard ${state(i)}">
        <div class="stepnum">${done[i] ? "✓" : s.n}</div>
        <div class="steptag">${done[i] ? "Done" : i === next ? "Next step" : "Step " + s.n}</div>
        <h3>${s.title}</h3>
        <p>${s.text}</p>
        <div class="stepdetail">${s.detail}</div>
        <a class="btn sm ${i === next ? "pri" : "soft"}" href="${s.href}">${s.btn} →</a>
      </div>`).join("")}
  </div>
  <div class="tips">
    <div class="tip"><b>Found a job on Seek or LinkedIn?</b> On the Apply tab, use “Add a job from another site” and paste the description. You get the same tailored documents.</div>
    <div class="tip"><b>Already applied somewhere?</b> Add it on the Tracker so it's counted and never suggested again.</div>
    <div class="tip"><b>Read before you send.</b> Documents only use facts from your resume, but a quick read makes sure they sound like you.</div>
  </div>`;
}

/* ============ PROFILE ============ */
function profileStatus() {
  if (S.profileDirty) return `<span class="dot warn"></span>You have unsaved changes`;
  if (S.profileMsg) return `<span class="dot ok"></span>${esc(S.profileMsg)}`;
  return `<span class="dot ok"></span>Everything is saved`;
}
function refreshSaveBar() {
  const bar = $("#savebar"), msg = $("#savemsg");
  if (!bar || !msg) return;
  bar.classList.toggle("dirty", !!S.profileDirty);
  msg.innerHTML = profileStatus();
}
function viewProfile() {
  const p = S.profile, pr = p.prefs;
  const rs = [...S.resumes.values()].sort((a, b) => a.addedAt - b.addedAt);
  return `
  <h1>My profile</h1>
  <p class="sub">Your resumes and skills are the only source used to tailor applications. Nothing is made up that isn't here.</p>
  <div class="grid2">
    <div>
      <div class="card">
        <h2>Resumes</h2>
        <p class="hint">Upload one or more versions (PDF, Word .docx or text). Tailoring can draw on all of them.</p>
        <label class="drop" id="drop">
          <input type="file" id="resfile" accept=".pdf,.docx,.txt,.md,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/plain" multiple class="sr" data-change="upload">
          <b>Drop resume files here</b> or click to choose
        </label>
        <div id="upmsg" class="small muted" style="margin-top:8px"></div>
        <div class="list" style="margin-top:12px">
          ${rs.length ? rs.map((r) => `
            <div class="ritem">
              <div class="t"><b>${esc(r.name)}</b><span class="small muted">${(r.text || "").length.toLocaleString()} characters · added ${fmtDate(r.addedAt)}</span></div>
              <button class="btn sm ghost" data-act="previewResume" data-id="${esc(r.id)}">View</button>
              <button class="btn sm ghost" data-act="delResume" data-id="${esc(r.id)}" aria-label="Delete">Delete</button>
            </div>`).join("") : `<div class="muted small">No resumes yet.</div>`}
        </div>
        <details style="margin-top:12px"><summary class="small" style="cursor:pointer;color:var(--accent-text)">Or paste resume text</summary>
          <label class="f">Name</label><input type="text" id="pasteName" placeholder="e.g. Resume – general">
          <label class="f">Resume text</label><textarea id="pasteText" rows="8" placeholder="Paste the full text of your resume"></textarea>
          <div style="margin-top:8px"><button class="btn sm" data-act="pasteResume">Add resume</button></div>
        </details>
      </div>
      <div class="card">
        <h2>Skills &amp; extra details</h2>
        <p class="hint">Anything worth using that isn't in your resume: skills, tools, achievements, licences, work rights, availability.</p>
        <label class="f">Skills</label>
        <textarea data-pf="skills" rows="5" placeholder="e.g. Python, SQL, Power BI, stakeholder management, customer service, forklift licence">${esc(p.skills)}</textarea>
        <label class="f">Notes for tailoring</label>
        <textarea data-pf="notes" rows="4" placeholder="e.g. Full Australian work rights. Available immediately. Prefer hybrid. Current Working With Children Check.">${esc(p.notes)}</textarea>
      </div>
    </div>
    <div>
      <div class="card">
        <h2>What jobs to look for</h2>
        <p class="hint">Used by Find jobs. Separate job titles with commas; each one is searched separately.</p>
        <label class="f">Job titles to search</label>
        <div class="row"><input type="text" data-pf="prefs.roles" value="${esc(pr.roles)}" placeholder="e.g. Data Analyst, Business Analyst, Reporting Analyst">
          <button class="btn sm shrink" data-act="suggestRoles" ${claudeReady() && S.resumes.size ? "" : "disabled"}>Suggest from my resume</button></div>
        <div id="rolemsg" class="small muted" style="margin-top:4px"></div>
        <div class="row" style="margin-top:2px">
          <div><label class="f">Location</label><input type="text" data-pf="prefs.location" value="${esc(pr.location)}" placeholder="City and state, or remote"></div>
          <div><label class="f">Country</label><select data-pf="prefs.country">${COUNTRIES.map(([c, n]) => `<option value="${c}" ${pr.country === c ? "selected" : ""}>${n}</option>`).join("")}</select></div>
        </div>
        <label class="f">Job type</label>
        <select data-pf="prefs.jobType">${JOBTYPES.map(([v, n]) => `<option value="${v}" ${pr.jobType === v ? "selected" : ""}>${n}</option>`).join("")}</select>
      </div>
      <div class="card savebar ${S.profileDirty ? "dirty" : ""}" id="savebar">
        <div style="flex:1" class="small" id="savemsg">${profileStatus()}</div>
        <button class="btn pri" data-act="saveProfile">Save profile</button>
        <a class="btn" href="#find">Find jobs →</a>
      </div>
    </div>
  </div>`;
}

document.addEventListener("input", (e) => {
  const el = e.target.closest("[data-pf]"); if (!el) return;
  const path = el.dataset.pf.split(".");
  let o = S.profile; for (let i = 0; i < path.length - 1; i++) o = o[path[i]];
  o[path[path.length - 1]] = el.value;
  S.profileDirty = true; S.profileMsg = null;
  refreshSaveBar();
});
document.addEventListener("change", (e) => {
  const el = e.target;
  if (el.matches("select[data-pf]")) { el.dispatchEvent(new Event("input", { bubbles: true })); }
  const c = el.closest("[data-change]"); if (c && CHANGE[c.dataset.change]) CHANGE[c.dataset.change](c, e);
});

async function extractText(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".pdf") || file.type === "application/pdf") {
    await loadScript("vendor/pdf.min.js"); await loadScript("vendor/pdf.worker.min.js");
    const lib = window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";
    const pdf = await lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false }).promise;
    let out = "";
    for (let i = 1; i <= pdf.numPages; i++) {
      const tc = await (await pdf.getPage(i)).getTextContent();
      out += tc.items.map((it) => (it.str || "") + (it.hasEOL ? "\n" : "")).join("") + "\n\n";
    }
    return out;
  }
  if (name.endsWith(".docx")) {
    await loadScript("vendor/mammoth.browser.min.js");
    const r = await window.mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
    return r.value;
  }
  if (name.endsWith(".doc")) throw new Error("Old .doc files can't be read. Save it as .docx or PDF and upload again.");
  return await file.text();
}
const tidyText = (t) => String(t || "").replace(/\r/g, "").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim().slice(0, 40000);

async function addResumeFiles(files) {
  const msg = $("#upmsg");
  for (const f of files) {
    if (S.resumes.size >= 6) { toast("You can keep up to 6 resumes. Delete one to add another."); break; }
    if (msg) msg.textContent = "Reading " + f.name + "…";
    try {
      const text = tidyText(await extractText(f));
      if (text.length < 200) throw new Error("Couldn't find much text in " + f.name + ". If it's a scanned PDF, paste the text instead.");
      await saveResume({ id: "r-" + now().toString(36) + Math.random().toString(36).slice(2, 6), name: f.name.replace(/\.[^.]+$/, ""), text, addedAt: now() });
      S.profileMsg = `Resume “${f.name.replace(/\.[^.]+$/, "")}” added and saved`; refreshSaveBar();
      toast("Added and saved " + f.name);
    } catch (e) { toast(e.message || "Couldn't read " + f.name, 6000); }
  }
  if ($("#upmsg")) $("#upmsg").textContent = "";
}
document.addEventListener("dragover", (e) => { const d = e.target.closest && e.target.closest("#drop"); if (d) { e.preventDefault(); d.classList.add("over"); } });
document.addEventListener("dragleave", (e) => { const d = e.target.closest && e.target.closest("#drop"); if (d) d.classList.remove("over"); });
document.addEventListener("drop", (e) => { const d = e.target.closest && e.target.closest("#drop"); if (d) { e.preventDefault(); d.classList.remove("over"); addResumeFiles([...e.dataTransfer.files]); } });

/* ============ Claude helpers ============ */
function sampleErr(e) {
  switch (e && e.code) {
    case "not_granted": case "sampling_disabled": case "not_declared": case "capability_disabled": case "capability_removed":
      return "This page isn't allowed to use Claude in this view, so AI features are off.";
    case "rate_limited": return "You've hit a usage limit. Wait a bit and try again.";
    case "session_expired": return "Your Claude session expired. Sign in again and retry.";
    case "refused": return "Claude declined this request. Try rewording your instructions.";
    case "invalid_json": case "empty_completion": return "Claude's answer came back in the wrong format. Try again.";
    case "prompt_too_large": return "Too much text to send at once. Remove a resume or shorten the job description.";
    case "cancelled": return "Stopped.";
    case "no_api_key": return "Claude isn't set up yet. Add ANTHROPIC_API_KEY to the .env file and restart the server.";
    case "invalid_key": return "The Claude API rejected your API key. Check ANTHROPIC_API_KEY in .env.";
    case "bad_model": return e.message;
    case "bad_request": return "The Claude API refused the request: " + (e.message || "");
    case "offline": return e.message;
    case "upstream_error": return (e.message || "The Claude API had a problem.") + " Try again.";
    default: return "Something went wrong talking to Claude. Try again.";
  }
}
function candidateMaterial(baseId, maxEach = 14000) {
  const rs = [...S.resumes.values()].sort((a, b) => a.addedAt - b.addedAt);
  const use = baseId && baseId !== "all" ? rs.filter((r) => r.id === baseId) : rs;
  let out = use.map((r, i) => `--- RESUME ${i + 1}: "${r.name}" ---\n${(r.text || "").slice(0, maxEach)}`).join("\n\n");
  if (S.profile.skills) out += `\n\n--- EXTRA SKILLS ---\n${S.profile.skills}`;
  if (S.profile.notes) out += `\n\n--- CANDIDATE NOTES ---\n${S.profile.notes}`;
  return out;
}

/* ============ setup & job search helpers ============ */
const claudeReady = () => !!(S.sample && S.status && S.status.claude);
const jobsReady = () => !!(S.board && S.status && S.status.jobs);
function boardErr(e) {
  switch (e && e.code) {
    case "no_keys": return "Job search isn't set up yet. Add your free Adzuna app ID and key to the .env file, then restart the server (see README).";
    case "bad_keys": return "Adzuna rejected your app ID or key. Check ADZUNA_APP_ID and ADZUNA_APP_KEY in .env, then restart.";
    case "bad_country": return e.message;
    case "rate_limited": return "Adzuna's rate limit was reached. Wait a minute and try again.";
    case "offline": return e.message;
    default: return "Job search failed. " + (e && e.message ? e.message : "Try again.");
  }
}
const postedTs = (p) => { const t = Date.parse(p || ""); return isNaN(t) ? null : t; };

/* tracked-job matching (so searches skip anything already in the tracker) */
function trackedIndex() {
  const keys = new Set(), loose = new Set(), urls = new Set();
  for (const j of S.jobs.values()) {
    keys.add(j.key);
    loose.add(norm(j.title) + "|" + (normCo(j.company).split(" ")[0] || ""));
    if (j.url) urls.add(j.url);
  }
  return { has: (r) => keys.has(keyFor(r.title, r.company)) || loose.has(norm(r.title) + "|" + (normCo(r.company).split(" ")[0] || "")) || (r.url && urls.has(r.url)) };
}

/* ============ FIND ============ */
function viewFind() {
  const F = S.find, pr = S.profile.prefs;
  const noResume = !S.resumes.size;
  const res = F.results;
  return `
  <h1>Find new jobs</h1>
  <p class="sub">Searches job listings from many sites (through Adzuna) for your target roles, skips anything already in your tracker, and ranks the rest against your resume.</p>
  ${noResume ? `<div class="notice acc"><b>Add your resume first.</b> Matching and tailoring use it. <a href="#profile">Go to My profile →</a></div>` : ""}
  ${!jobsReady() ? `<div class="notice warn"><b>Job search isn't set up yet.</b> Add your free Adzuna keys to <code>.env</code> and restart the server (steps in the README). Meanwhile you can add jobs by hand on the <a href="#apply">Apply</a> tab.</div>` : ""}
  <div class="card">
    <div class="row">
      <div style="flex:2"><label class="f">Job titles (comma separated)</label><input type="text" data-pf="prefs.roles" value="${esc(pr.roles)}" placeholder="e.g. Data Analyst, Business Analyst"></div>
      <div><label class="f">Location</label><input type="text" data-pf="prefs.location" value="${esc(pr.location)}" placeholder="City, or remote"></div>
      <div style="max-width:160px"><label class="f">Country</label><select data-pf="prefs.country">${COUNTRIES.map(([c, n]) => `<option value="${c}" ${pr.country === c ? "selected" : ""}>${n}</option>`).join("")}</select></div>
      <div style="max-width:150px"><label class="f">Type</label><select data-pf="prefs.jobType">${JOBTYPES.map(([v, n]) => `<option value="${v}" ${pr.jobType === v ? "selected" : ""}>${n}</option>`).join("")}</select></div>
      <div style="max-width:110px"><label class="f">Show me</label><select id="findCount" data-change="count">${[10, 15, 20, 25].map((n) => `<option ${F.count === n ? "selected" : ""}>${n}</option>`).join("")}</select></div>
      <div class="shrink"><button class="btn pri" data-act="find" ${F.busy || !jobsReady() ? "disabled" : ""}>${F.busy ? "Searching…" : "Find new jobs"}</button></div>
    </div>
    ${F.busy ? `<div class="progress"><div class="spin"></div><span id="findmsg">${esc(F.msg)}</span></div>` : ""}
    ${F.err ? `<div class="notice bad" style="margin:12px 0 0">${esc(F.err)}</div>` : ""}
  </div>
  ${res ? findResults() : ""}`;
}
function findResults() {
  const F = S.find, res = F.results;
  if (!res.length) return `<div class="card empty"><b>No new jobs found</b>${F.hidden ? `${plural(F.hidden, "job")} matched but ${F.hidden === 1 ? "is" : "are"} already in your tracker. ` : ""}Try other job titles or a wider location.</div>`;
  const sel = F.selected;
  return `
  <div class="toolbar">
    <div><b>${res.length} new ${res.length === 1 ? "job" : "jobs"}</b> <span class="muted">${F.hidden ? `· ${F.hidden} already tracked, hidden` : ""}${F.short ? ` · fewer than you asked for; add more job titles to widen the search` : ""}</span></div>
    <div class="sp"></div>
    <button class="btn sm" data-act="selAll">Select all</button><button class="btn sm ghost" data-act="selNone">Clear</button>
  </div>
  <div class="jobs">
    ${res.map((r, i) => {
      const on = sel.has(i), sc = r.score;
      const pt = postedTs(r.posted), old = pt && daysSince(pt) > 45;
      return `<div class="job ${on ? "sel" : ""}">
        <input type="checkbox" data-change="pick" data-i="${i}" ${on ? "checked" : ""} aria-label="Select ${esc(r.title)}">
        <div class="main">
          <a class="title" href="${esc(safeUrl(r.url))}" target="_blank" rel="noopener">${esc(r.title)}</a>
          <div class="meta">${esc(r.company)} · ${esc(r.location)}${r.jobType ? " · " + esc(r.jobType) : ""}${r.pay ? " · " + esc(r.pay) : ""}</div>
          <div class="why">${r.reason ? esc(r.reason) + " · " : ""}<span class="${old ? "" : "muted"}">${pt ? "Posted " + ago(pt) : esc(r.posted)}</span>${old ? ` <span class="pill warn">May be closed</span>` : ""}</div>
        </div>
        ${sc != null ? `<div class="score ${sc >= 75 ? "hi" : sc >= 55 ? "md" : ""}" data-tip="Match with your resume, based on title and company">${sc}<small>match</small></div>` : ""}
        <button class="btn sm ghost" data-act="hideResult" data-i="${i}" data-tip="Not interested: never show this job again">Hide</button>
      </div>`;
    }).join("")}
  </div>
  <div class="stickybar">
    <div style="flex:1"><b>${sel.size}</b> selected</div>
    <button class="btn pri" data-act="addSelected" ${sel.size && !F.adding ? "" : "disabled"}>${F.adding ? "Adding…" : "Add to my apply list →"}</button>
  </div>`;
}

async function runFind() {
  const F = S.find, pr = S.profile.prefs;
  if (S.profileDirty) saveProfile();
  const roles = [...new Set(String(pr.roles || "").split(/[,;\n]/).map((s) => s.trim()).filter(Boolean))].slice(0, 6);
  if (!roles.length) { F.err = "Add at least one job title to search for (or use Suggest from my resume on My profile)."; renderView(); return; }
  if (!pr.location.trim()) { F.err = "Add a location (a city, or 'remote')."; renderView(); return; }
  Object.assign(F, { busy: true, err: "", results: null, selected: new Set(), msg: "Searching job listings…" });
  renderView();
  const setMsg = (m) => { F.msg = m; const el = $("#findmsg"); if (el) el.textContent = m; };
  try {
    const all = [];
    let lastErr = null, okCount = 0;
    for (let i = 0; i < roles.length; i++) {
      setMsg(`Searching for “${roles[i]}” (${i + 1} of ${roles.length})…`);
      const input = { what: roles[i], where: pr.location.trim(), country: (pr.country || "AU").toLowerCase() };
      if (pr.jobType) input.jobType = pr.jobType;
      try { all.push(...(await S.board.search(input))); okCount++; }
      catch (e) {
        lastErr = e;
        if (["no_keys", "bad_keys", "bad_country", "offline"].includes(e.code)) break;
      }
    }
    if (!okCount && lastErr) throw lastErr;
    // dedupe within results, then drop anything already tracked
    const seen = new Set(), idx = trackedIndex(); let hidden = 0;
    const fresh = [];
    for (const r of all) {
      const k = norm(r.title) + "|" + (normCo(r.company).split(" ")[0] || "") + "|" + norm(r.location).split(" ")[0];
      if (seen.has(k)) continue; seen.add(k);
      if (idx.has(r)) { hidden++; continue; }
      fresh.push(r);
    }
    // newest first as a base order
    fresh.sort((a, b) => (postedTs(b.posted) || 0) - (postedTs(a.posted) || 0));
    if (fresh.length && claudeReady() && S.resumes.size) {
      setMsg(`Ranking ${fresh.length} jobs against your resume…`);
      try {
        const scores = await S.sample.json(rankPrompt(fresh), { modelTier: "quick" });
        if (Array.isArray(scores)) scores.forEach((s) => { const r = fresh[Number(s.i)]; if (r) { r.score = Math.max(0, Math.min(100, Math.round(Number(s.score) || 0))); r.reason = String(s.reason || "").slice(0, 140); } });
        fresh.sort((a, b) => (b.score ?? -1) - (a.score ?? -1));
      } catch (e) { if (e.code !== "cancelled") toast("Couldn't rank results: " + sampleErr(e), 5000); }
    }
    F.results = fresh.slice(0, F.count);
    F.short = fresh.length < F.count && fresh.length > 0;
    F.hidden = hidden;
  } catch (e) {
    F.err = boardErr(e);
  } finally {
    F.busy = false; renderView();
  }
}
function rankPrompt(list) {
  const rs = [...S.resumes.values()].sort((a, b) => a.addedAt - b.addedAt);
  const cv = (rs[0] && rs[0].text || "").slice(0, 5000);
  return `You screen job listings for a job seeker. Score how well each job fits the candidate from 0 to 100, using the job title, company, location and type against the candidate's experience, seniority and skills. Penalise jobs clearly outside their field or far above/below their level.

CANDIDATE TARGET ROLES: ${S.profile.prefs.roles}
CANDIDATE SKILLS: ${(S.profile.skills || "").slice(0, 1500)}
CANDIDATE RESUME (excerpt):
${cv}

JOBS:
${list.map((r, i) => `${i} | ${r.title} | ${r.company} | ${r.location} | ${r.jobType} | posted ${r.posted}`).join("\n")}

Reply with only a JSON array with one object per job, e.g. [{"i":0,"score":82,"reason":"Strong SQL and reporting overlap; right seniority"}]. Keep each reason under 14 words and specific to this candidate.`;
}

async function addSelected() {
  const F = S.find; if (!F.results) return;
  const picks = [...F.selected].sort((a, b) => a - b).map((i) => F.results[i]).filter(Boolean);
  if (!picks.length) return;
  F.adding = true; renderView();
  const added = [];
  for (const r of picks) {
    const key = keyFor(r.title, r.company);
    if (S.jobs.has(key)) continue;
    const job = { key, title: r.title, company: r.company, location: r.location, jobType: r.jobType, pay: r.pay, posted: r.posted, url: safeUrl(r.url), boardId: r.id, source: "adzuna",
      status: "shortlisted", match: r.score != null ? { score: r.score, reason: r.reason || "" } : null, description: r.snippet || "", descIsSnippet: !!r.snippet, addedAt: now(), history: [{ s: "shortlisted", at: now() }], notes: "" };
    saveJob(job); added.push(job);
  }
  F.results = F.results.filter((_, i) => !F.selected.has(i));
  F.selected = new Set(); F.adding = false;
  renderView();
  toast(`Added ${plural(added.length, "job")} to your apply list.`, 4000);
}

/* ============ APPLY ============ */
function viewApply() {
  const list = [...S.jobs.values()].filter((j) => TO_APPLY.has(j.status)).sort((a, b) => (b.match?.score ?? -1) - (a.match?.score ?? -1) || b.addedAt - a.addedAt);
  const asking = list.filter((j) => j.openedAt);
  return `
  <h1>Apply</h1>
  <p class="sub">Pick a job, get a resume and cover letter tailored to it, download them, then apply on the job site.</p>
  ${asking.map((j) => `<div class="notice acc" style="display:flex;align-items:center;gap:10px;flex-wrap:wrap"><div style="flex:1">You opened <b>${esc(j.title)}</b> at ${esc(j.company)} ${ago(j.openedAt)}. Did you apply?</div>
     <button class="btn sm good" data-act="markApplied" data-key="${esc(j.key)}">Yes, I applied</button><button class="btn sm" data-act="notYet" data-key="${esc(j.key)}">Not yet</button></div>`).join("")}
  <div class="toolbar"><b>${plural(list.length, "job")} to apply for</b><div class="sp"></div>
    <button class="btn sm" data-act="manualAdd" data-status="shortlisted">+ Add a job from another site</button>
    <a class="btn sm pri" href="#find">Find more jobs</a></div>
  ${list.length ? `<div class="jobs">${list.map(applyCard).join("")}</div>` :
    `<div class="card empty"><b>Nothing waiting to apply</b>Select jobs on <a href="#find">Find jobs</a>, or add one you found on Seek, LinkedIn or a company site.</div>`}`;
}
function applyCard(j) {
  const sc = j.match?.score, ready = j.status === "ready" && j.tailored;
  return `<div class="job">
    ${sc != null ? `<div class="score ${sc >= 75 ? "hi" : sc >= 55 ? "md" : ""}">${sc}<small>match</small></div>` : ""}
    <div class="main">
      ${j.url ? `<a class="title" href="${esc(j.url)}" target="_blank" rel="noopener">${esc(j.title)}</a>` : `<span class="title">${esc(j.title)}</span>`}
      <div class="meta">${esc(j.company)}${j.location ? " · " + esc(j.location) : ""}${j.jobType ? " · " + esc(j.jobType) : ""}${j.pay ? " · " + esc(j.pay) : ""}</div>
      <div class="why">${ready ? `<span class="pill good">✓ Resume &amp; cover letter ready</span> ` : `<span class="pill">No documents yet</span> `}${!j.description ? `<span class="pill warn">No job description yet</span> ` : j.descIsSnippet ? `<span class="pill warn">Summary only: paste the full description</span> ` : ""}${j.match?.reason ? esc(j.match.reason) : ""}</div>
    </div>
    <div class="acts">
      <button class="btn sm ${ready ? "" : "pri"}" data-act="openWs" data-key="${esc(j.key)}">${ready ? "Open documents" : "Tailor resume &amp; letter"}</button>
      ${j.url ? `<a class="btn sm" href="${esc(j.url)}" target="_blank" rel="noopener" data-act="opened" data-key="${esc(j.key)}">Apply on site ↗</a>` : ""}
      <button class="btn sm good" data-act="markApplied" data-key="${esc(j.key)}">Mark applied</button>
      <button class="btn sm ghost" data-act="dismiss" data-key="${esc(j.key)}">Not interested</button>
    </div>
  </div>`;
}

/* ---------- tailoring workspace ---------- */
function openWs(key) {
  const j = S.jobs.get(key); if (!j) return;
  S.ws = { key, busy: false, tab: "resume", err: "", base: "all", extra: "", revise: "", desc: j.description || "", descDirty: false, ctl: null };
  renderWs();
  loadDocLibs().catch(() => {});
}
function closeWs() {
  if (S.ws && S.ws.ctl) S.ws.ctl.abort();
  S.ws = null; $("#modal").hidden = true; $("#modal").innerHTML = "";
  scheduleRender();
}
function renderWs() {
  const W = S.ws; if (!W) return;
  const j = S.jobs.get(W.key); if (!j) { closeWs(); return; }
  const t = j.tailored, rs = [...S.resumes.values()].sort((a, b) => a.addedAt - b.addedAt);
  const m = $("#modal");
  m.innerHTML = `<div class="mbox wide" role="dialog" aria-label="Tailor documents">
    <div class="ws-head"><div class="t">
      <h3>${esc(j.title)}</h3><div class="muted">${esc(j.company)}${j.location ? " · " + esc(j.location) : ""}${j.url ? ` · <a href="${esc(j.url)}" target="_blank" rel="noopener">View posting ↗</a>` : ""}</div></div>
      <button class="btn sm ghost" data-act="closeWs" aria-label="Close">✕ Close</button></div>
    <div class="ws-body">
      <div class="ws-left">
        <p class="step">Step 1 · Job description</p>
        ${j.descIsSnippet && !W.descDirty ? `<div class="notice warn small" style="margin-bottom:8px">This is only a short summary from the job board. For a much better result, ${j.url ? `<a href="${esc(j.url)}" target="_blank" rel="noopener">open the posting ↗</a>` : "open the posting"}, copy the full description and paste it here.</div>` : ""}
        <textarea id="wsDesc" placeholder="Paste the full job description here (duties, requirements, selection criteria)…">${esc(W.desc)}</textarea>
        <div class="small muted" style="margin-top:4px">${W.desc ? `${W.desc.length.toLocaleString()} characters` : "Needed for a good tailored resume."}</div>
        <p class="step" style="margin-top:18px">Step 2 · Tailor</p>
        ${!rs.length ? `<div class="notice warn">Add a resume on <a href="#profile" data-act="closeWs">My profile</a> first.</div>` : `
        <label class="f">Base it on</label>
        <select id="wsBase">${rs.length > 1 ? `<option value="all" ${W.base === "all" ? "selected" : ""}>All my resumes (pick the best of each)</option>` : ""}${rs.map((r) => `<option value="${esc(r.id)}" ${W.base === r.id ? "selected" : ""}>${esc(r.name)}</option>`).join("")}</select>
        <label class="f">Anything to emphasise? <span class="muted">(optional)</span></label>
        <input type="text" id="wsExtra" value="${esc(W.extra)}" placeholder="e.g. focus on my customer service; keep it to one page">`}
        <div style="margin-top:12px;display:flex;gap:8px">
          ${W.busy ? `<button class="btn" data-act="wsStop">Stop</button>` : `<button class="btn pri" data-act="wsGenerate" ${rs.length && claudeReady() ? "" : "disabled"}>${t ? "Regenerate" : "Generate resume + cover letter"}</button>`}
        </div>
        ${!claudeReady() ? `<div class="notice warn small" style="margin-top:10px">Tailoring needs your Anthropic API key. Add <code>ANTHROPIC_API_KEY</code> to <code>.env</code> and restart the server.</div>` : ""}
      </div>
      <div class="ws-right">
        ${W.err ? `<div class="notice bad">${esc(W.err)}</div>` : ""}
        ${W.busy ? `<div class="progress"><div class="spin"></div><span id="wsProg">${esc(W.prog || "Claude is reading the job and your resume…")}</span></div><p class="small muted">This usually takes 30–90 seconds. Only facts from your resumes are used.</p>` : ""}
        ${!W.busy && !t ? `<div class="empty"><b>Your tailored documents will appear here</b>Check the job description on the left, then click Generate.</div>` : ""}
        ${!W.busy && t ? wsResult(j) : ""}
      </div>
    </div>
    <div class="applybar">
      <p class="step" style="margin:0">Step 3 · Apply</p><div class="sp" style="flex:1"></div>
      ${j.url ? `<a class="btn pri" href="${esc(j.url)}" target="_blank" rel="noopener" data-act="opened" data-key="${esc(j.key)}">Open job posting &amp; apply ↗</a>` : ""}
      <button class="btn good" data-act="markApplied" data-key="${esc(j.key)}">I've applied ✓</button>
    </div>
  </div>`;
  m.hidden = false;
}
function wsResult(j) {
  const t = j.tailored, W = S.ws, fit = t.fit || {};
  const r = JTDocs.normalizeResume(t.resume), c = JTDocs.normalizeLetter(t.coverLetter, t.resume);
  const li = (a) => (a && a.length ? `<ul>${a.slice(0, 8).map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : `<div class="small muted">None noted</div>`);
  const paper = W.tab === "resume" ? `
    <div class="nm">${esc(r.name)}</div>${r.headline ? `<div class="hl">${esc(r.headline)}</div>` : ""}<div class="ct">${esc(r.contact.join("  |  "))}</div>
    ${r.summary ? `<h5>Professional Summary</h5><p>${esc(r.summary)}</p>` : ""}
    ${r.skills.length ? `<h5>Key Skills</h5>${r.skills.map((g) => `<div>${g.group ? `<b>${esc(g.group)}:</b> ` : ""}${esc(g.items.join(", "))}</div>`).join("")}` : ""}
    ${r.experience.length ? `<h5>Professional Experience</h5>${r.experience.map((e) => `<div class="rl"><span><b>${esc(e.title)}</b>${e.company ? " · " + esc(e.company) : ""}</span><span>${esc(e.dates)}</span></div>${e.location ? `<div class="loc">${esc(e.location)}</div>` : ""}<ul>${e.bullets.map((b) => `<li>${esc(b)}</li>`).join("")}</ul>`).join("")}` : ""}
    ${r.projects.length ? `<h5>Projects</h5>${r.projects.map((p) => `<div class="rl"><b>${esc(p.name)}</b></div>${p.description ? `<div>${esc(p.description)}</div>` : ""}<ul>${p.bullets.map((b) => `<li>${esc(b)}</li>`).join("")}</ul>`).join("")}` : ""}
    ${r.education.length ? `<h5>Education</h5>${r.education.map((e) => `<div class="rl"><span><b>${esc(e.qualification)}</b>${e.institution ? " · " + esc(e.institution) : ""}</span><span>${esc(e.dates)}</span></div>${e.details ? `<div>${esc(e.details)}</div>` : ""}`).join("")}` : ""}
    ${r.certifications.length ? `<h5>Certifications &amp; Licences</h5><ul>${r.certifications.map((b) => `<li>${esc(b)}</li>`).join("")}</ul>` : ""}
    ${r.additional.length ? `<h5>Additional Information</h5><ul>${r.additional.map((b) => `<li>${esc(b)}</li>`).join("")}</ul>` : ""}` : `
    <div style="font-size:18px;font-weight:700;color:#1f4e79">${esc(c.name)}</div><div class="small" style="color:#444;border-bottom:1px solid #9aa5b1;padding-bottom:8px;margin-bottom:14px">${esc(c.contact.join("  |  "))}</div>
    ${c.date ? `<p>${esc(c.date)}</p>` : ""}<p>${c.recipient.map(esc).join("<br>")}</p>${c.subject ? `<p><b>${esc(c.subject)}</b></p>` : ""}
    <p>${esc(c.greeting)}</p>${c.paragraphs.map((p) => `<p>${esc(p)}</p>`).join("")}<p>${esc(c.closing)}</p><p><b>${esc(c.name)}</b></p>`;
  const kind = W.tab === "resume" ? "resume" : "letter";
  return `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
      ${fit.score != null ? `<div class="score ${fit.score >= 75 ? "hi" : fit.score >= 55 ? "md" : ""}">${esc(fit.score)}<small>fit</small></div>` : ""}
      <div style="flex:1" class="small">${esc(fit.summary || "")}<div class="muted">Generated ${ago(t.at)}</div></div>
    </div>
    <div class="fit">
      <div class="box"><h4>✓ You match</h4>${li(fit.matched)}</div>
      <div class="box"><h4>△ Gaps to address</h4>${li(fit.gaps)}</div>
    </div>
    ${fit.checklist && fit.checklist.length ? `<div class="notice warn" style="margin-bottom:14px"><b>Before you apply:</b><ul style="margin:4px 0 0;padding-left:18px">${fit.checklist.slice(0, 8).map((x) => `<li>${esc(x)}</li>`).join("")}</ul></div>` : ""}
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <div class="seg"><button class="${W.tab === "resume" ? "on" : ""}" data-act="wsTab" data-tab="resume">Resume</button><button class="${W.tab === "letter" ? "on" : ""}" data-act="wsTab" data-tab="letter">Cover letter</button></div>
      <div class="sp" style="flex:1"></div>
      <button class="btn sm pri" data-act="dl" data-kind="${kind}" data-fmt="docx">↓ Word (.docx)</button>
      <button class="btn sm" data-act="dl" data-kind="${kind}" data-fmt="pdf">↓ PDF</button>
      <button class="btn sm ghost" data-act="copyDoc" data-kind="${kind}">Copy text</button>
    </div>
    <div class="paper">${paper}</div>
    <div class="row" style="margin-top:12px">
      <input type="text" id="wsRevise" value="${esc(W.revise)}" placeholder="Ask for changes, e.g. shorten the cover letter, mention my Excel skills more">
      <button class="btn shrink" data-act="wsRevise" ${claudeReady() ? "" : "disabled"}>Revise</button>
    </div>`;
}

const RESUME_SHAPE = `{
  "resume": {
    "name": "Full Name",
    "headline": "short target title line, e.g. Data Analyst | SQL · Power BI · Stakeholder reporting",
    "contact": ["email", "phone", "City STATE", "linkedin or portfolio URL if in source"],
    "summary": "3-4 sentence professional summary aimed at this job",
    "skills": [{"group": "Technical", "items": ["..."]}, {"group": "Professional", "items": ["..."]}],
    "experience": [{"title": "", "company": "", "location": "", "dates": "", "bullets": ["..."]}],
    "projects": [{"name": "", "description": "", "bullets": []}],
    "education": [{"qualification": "", "institution": "", "dates": "", "details": ""}],
    "certifications": ["..."],
    "additional": ["work rights, licences, checks, languages, availability - only if in source"]
  },
  "coverLetter": {
    "date": "", "recipient": ["Hiring Manager or named contact", "Company", "City"],
    "subject": "Re: Job Title (reference number if given)",
    "greeting": "Dear ...,", "paragraphs": ["...", "...", "..."], "closing": "Kind regards,"
  },
  "fit": {"score": 0, "summary": "one sentence on overall fit", "matched": ["key requirement the candidate clearly meets"], "gaps": ["requirement not clearly shown, with a tip"], "checklist": ["things the ad asks for when applying: selection criteria, checks, documents, closing date"]}
}`;
function rules(country) {
  const sp = { US: "American", CA: "Canadian", GB: "British", IE: "British", AU: "Australian", NZ: "New Zealand" }[country] || "British";
  const today = new Date().toLocaleDateString(country === "US" ? "en-US" : "en-AU", { day: "numeric", month: "long", year: "numeric" });
  return `RULES (follow strictly):
- Use ONLY facts found in the candidate material. Never invent or upgrade employers, job titles, dates, degrees, certifications, licences, tools, numbers or achievements. If a metric isn't in the source, don't add one. If unsure, leave it out.
- You may reorder, rewrite and trim to highlight what this job asks for, choose which skills to show, and use the ad's own keywords wherever the candidate genuinely has that skill or experience (this helps applicant tracking systems).
- Keep every real job in "experience" in reverse date order with dates exactly as in the source; older or less relevant roles can be cut to 1-2 bullets.
- Bullets: one sentence each, start with a strong verb, show impact. 3-6 bullets for recent relevant roles. The resume should fit on 1-2 pages.
- Contact details: copy from the source resume. Never invent them. Empty arrays/strings for sections with no source content.
- ${sp} spelling and conventions. Plain, confident, human tone. Avoid clichés ("results-driven", "passionate", "proven track record", "leverage", "dynamic", "synergy") and avoid em dashes.
- Cover letter: 250-380 words, 3-5 paragraphs, written for this company and role. Open with the role and a genuine reason for interest; give 2-3 concrete examples from the candidate's background matched to the key requirements; close with availability and a call to action. Don't restate the resume line by line. No placeholders like [Company]. Use the contact person's name if the ad gives one, else "Dear Hiring Manager,". Date: ${today}.
- "fit.checklist": list what the ad asks applicants to do or have (selection criteria responses, police/working-with-children checks, licences, documents, closing date). Empty if none.`;
}
function tailorPrompt(j, W) {
  return `You are an expert resume writer. Tailor the candidate's resume and write a cover letter for the job below.

${rules(S.profile.prefs.country)}
${W.extra ? `\nCANDIDATE'S REQUEST: ${W.extra}\n` : ""}
=== JOB ===
Title: ${j.title}
Company: ${j.company}
Location: ${j.location || ""}
Type: ${j.jobType || ""}
${j.pay ? "Pay: " + j.pay + "\n" : ""}Description:
${(W.desc || "(No description available. Tailor to the job title and typical requirements, and say so in fit.summary.)").slice(0, 15000)}

=== CANDIDATE MATERIAL ===
${candidateMaterial(W.base)}

Reply with only JSON in exactly this shape:
${RESUME_SHAPE}`;
}
function revisePrompt(j, W, prev) {
  return `You are an expert resume writer. Below is a tailored resume and cover letter (JSON) for a job, the job description, and the candidate's source material. Apply the candidate's requested change and return the complete updated JSON in the same shape. Change only what the request needs.

CANDIDATE'S REQUEST: ${W.revise}

${rules(S.profile.prefs.country)}

=== JOB ===
${j.title} at ${j.company}
${(W.desc || "").slice(0, 12000)}

=== CANDIDATE MATERIAL (source of truth) ===
${candidateMaterial(W.base, 12000)}

=== CURRENT DOCUMENTS (JSON) ===
${JSON.stringify({ resume: prev.resume, coverLetter: prev.coverLetter, fit: prev.fit })}

Reply with only the full updated JSON object.`;
}
async function wsRun(mode) {
  const W = S.ws; if (!W || W.busy || !S.sample) return;
  const j = S.jobs.get(W.key);
  if (W.descDirty) { const nj = clone(j); nj.description = W.desc; nj.descIsSnippet = false; saveJob(nj); W.descDirty = false; }
  if (mode === "revise" && !W.revise.trim()) { toast("Type what you'd like changed first."); return; }
  W.busy = true; W.err = ""; W.prog = ""; W.ctl = new AbortController(); renderWs();
  try {
    const prompt = mode === "revise" ? revisePrompt(j, W, j.tailored) : tailorPrompt(j, W);
    const out = await S.sample.json(prompt, {
      signal: W.ctl.signal, cache: false,
      onText: ({ text }) => { const el = $("#wsProg"); if (el) el.textContent = `Writing your documents… (${Math.round(text.length / 100) / 10}k characters)`; },
    });
    if (!out || !out.resume || !out.coverLetter) throw { code: "invalid_json" };
    const cur = S.jobs.get(W.key); if (!cur) return;
    const nj = clone(cur);
    nj.tailored = { resume: out.resume, coverLetter: out.coverLetter, fit: out.fit || {}, at: now(), base: W.base };
    if (nj.status === "shortlisted") { nj.status = "ready"; nj.history = (nj.history || []).concat([{ s: "ready", at: now() }]); }
    if (out.fit && out.fit.score != null) nj.match = Object.assign({}, nj.match || {}, { score: Math.round(Number(out.fit.score) || 0), reason: (out.fit.summary || nj.match?.reason || "").slice(0, 160) });
    if (mode === "revise") W.revise = "";
    saveJob(nj);
  } catch (e) {
    if (S.ws === W && e.code !== "cancelled") W.err = sampleErr(e);
  } finally {
    if (S.ws === W) { W.busy = false; W.ctl = null; renderWs(); }
  }
}

function docFilename(j, kind, fmt) {
  const nm = (j.tailored && j.tailored.resume && j.tailored.resume.name) || "";
  const parts = [nm, kind === "resume" ? "Resume" : "Cover Letter", j.company].filter(Boolean).join(" - ");
  return parts.replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ").slice(0, 120) + "." + fmt;
}
async function downloadDoc(key, kind, fmt) {
  const j = S.jobs.get(key); if (!j || !j.tailored) return;
  if (!S.downloads) { toast("Downloads aren't available in this view. Use Copy text instead.", 5000); return; }
  try {
    await loadDocLibs();
    const t = j.tailored, opts = { country: S.profile.prefs.country };
    let blob;
    if (fmt === "docx") {
      const d = kind === "resume" ? JTDocs.resumeDocx(window.docx, t.resume, opts) : JTDocs.letterDocx(window.docx, t.coverLetter, t.resume, opts);
      blob = await window.docx.Packer.toBlob(d);
    } else {
      const p = kind === "resume" ? JTDocs.resumePdf(window.jspdf.jsPDF, t.resume, opts) : JTDocs.letterPdf(window.jspdf.jsPDF, t.coverLetter, t.resume, opts);
      blob = p.output("blob");
    }
    const filename = docFilename(j, kind, fmt);
    await S.downloads.save({ filename, data: blob });
    toast("Downloaded " + filename);
  } catch (e) {
    if (e && e.code === "declined") return;
    if (e && e.code === "rate_limited") { toast("Another download is waiting for your OK. Finish that first."); return; }
    console.warn(e); toast("Couldn't create the file. Try again, or use Copy text.", 5000);
  }
}

/* ---------- did-you-apply ---------- */
function ask(html, buttons) {
  const a = $("#ask");
  a.innerHTML = `<div class="mbox" role="dialog">${html}<div class="mfoot">${buttons.map((b, i) => `<button class="btn ${b.cls || ""}" data-ask="${i}">${b.label}</button>`).join("")}</div></div>`;
  a.hidden = false;
  a.onclick = (e) => {
    const b = e.target.closest("[data-ask]");
    if (!b && e.target !== a) return;
    a.hidden = true; a.innerHTML = ""; a.onclick = null;
    if (b) { const fn = buttons[+b.dataset.ask].fn; if (fn) fn(); }
  };
}
function askApplied(key) {
  const j = S.jobs.get(key); if (!j || !TO_APPLY.has(j.status)) return;
  ask(`<h3>Did you apply?</h3><p class="muted" style="margin:0"><b style="color:var(--text)">${esc(j.title)}</b> at ${esc(j.company)}</p>`, [
    { label: "Not going to apply", cls: "ghost", fn: () => { setStatus(key, "dismissed"); if (S.ws && S.ws.key === key) closeWs(); } },
    { label: "Not yet", fn: () => clearOpened(key) },
    { label: "Yes, I applied", cls: "good", fn: () => { setStatus(key, "applied"); toast("Marked as applied. It's on your Tracker."); if (S.ws && S.ws.key === key) closeWs(); } },
  ]);
}
function clearOpened(key) { const j = S.jobs.get(key); if (j && j.openedAt) { const nj = clone(j); nj.openedAt = null; saveJob(nj); } }
function maybeAsk() {
  const p = S.pendingAsk; if (!p || now() - p.at < 3000 || document.visibilityState !== "visible") return;
  S.pendingAsk = null; askApplied(p.key);
}
document.addEventListener("visibilitychange", maybeAsk);
window.addEventListener("focus", maybeAsk);

/* ---------- manual add ---------- */
function manualAdd(status) {
  const m = $("#modal");
  m.innerHTML = `<div class="mbox" role="dialog"><h3>Add a job</h3><p class="muted small" style="margin:0 0 6px">For jobs from Seek, LinkedIn, company sites, or ones you already applied to, so they're tracked and never suggested again.</p>
    <form id="manForm">
    <div class="row"><div><label class="f">Job title *</label><input type="text" name="title" required></div><div><label class="f">Company *</label><input type="text" name="company" required></div></div>
    <div class="row"><div><label class="f">Location</label><input type="text" name="location"></div><div><label class="f">Link to posting</label><input type="url" name="url" placeholder="https://"></div></div>
    <div class="row"><div><label class="f">Status</label><select name="status">${STATUSES.map(([v, l]) => `<option value="${v}" ${v === status ? "selected" : ""}>${l}</option>`).join("")}</select></div>
      <div><label class="f">Date applied</label><input type="date" name="applied" value="${status === "applied" ? isoDate(now()) : ""}"></div></div>
    <label class="f">Job description <span class="muted">(paste it to get a tailored resume)</span></label><textarea name="description" rows="5"></textarea>
    <div class="mfoot"><button type="button" class="btn ghost" data-act="closeModal">Cancel</button><button class="btn pri" type="submit">Add job</button></div></form></div>`;
  m.hidden = false;
  $("#manForm").onsubmit = (e) => {
    e.preventDefault();
    const f = Object.fromEntries(new FormData(e.target).entries());
    const key = keyFor(f.title, f.company);
    if (S.jobs.has(key)) { toast("That job is already in your tracker."); return; }
    const st = f.status || "shortlisted";
    const appliedAt = f.applied ? new Date(f.applied + "T12:00:00").getTime() : APPLIED_SET.has(st) ? now() : null;
    saveJob({ key, title: f.title.trim(), company: f.company.trim(), location: (f.location || "").trim(), url: safeUrl(f.url), source: "manual", status: st,
      description: (f.description || "").trim().slice(0, 15000), addedAt: now(), appliedAt, respondedAt: HEARD_SET.has(st) ? now() : null, history: [{ s: st, at: now() }], notes: "", match: null });
    m.hidden = true; m.innerHTML = "";
    toast("Added " + f.title);
  };
}

/* ============ TRACKER ============ */
function stats() {
  const all = [...S.jobs.values()];
  const live = all.filter((j) => j.status !== "dismissed");
  const applied = all.filter((j) => APPLIED_SET.has(j.status));
  const heard = all.filter((j) => HEARD_SET.has(j.status));
  const interviews = all.filter((j) => j.status === "interview" || j.status === "offer" || (j.history || []).some((h) => h.s === "interview"));
  const offers = all.filter((j) => j.status === "offer");
  const waiting = all.filter((j) => j.status === "applied");
  const toApply = all.filter((j) => TO_APPLY.has(j.status));
  return { all, live, applied, heard, interviews, offers, waiting, toApply };
}
function weekStart(t) { const d = new Date(t); d.setHours(0, 0, 0, 0); const day = (d.getDay() + 6) % 7; d.setDate(d.getDate() - day); return d.getTime(); }
function viewTracker() {
  const st = stats(), T = S.tracker;
  const rate = st.applied.length ? Math.round((st.heard.length / st.applied.length) * 100) : null;
  // weekly applications, last 8 weeks
  const w0 = weekStart(now()), weeks = [];
  for (let i = 7; i >= 0; i--) weeks.push({ start: w0 - i * 7 * 86400000, n: 0 });
  st.applied.forEach((j) => { if (!j.appliedAt) return; const ws = weekStart(j.appliedAt); const w = weeks.find((x) => x.start === ws); if (w) w.n++; });
  const wmax = Math.max(1, ...weeks.map((w) => w.n));
  const scounts = STATUSES.filter(([s]) => s !== "dismissed").map(([s, l]) => ({ s, l, n: st.all.filter((j) => j.status === s).length }));
  const smax = Math.max(1, ...scounts.map((x) => x.n));
  const kpi = (l, v, s) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;

  // table rows
  const q = norm(T.q);
  let rows = st.all.filter((j) => {
    if (T.filter === "all") { if (j.status === "dismissed") return false; }
    else if (T.filter === "toapply") { if (!TO_APPLY.has(j.status)) return false; }
    else if (T.filter === "active") { if (!["applied", "response", "interview", "offer"].includes(j.status)) return false; }
    else if (j.status !== T.filter) return false;
    return !q || norm(j.title + " " + j.company + " " + (j.location || "") + " " + (j.notes || "")).includes(q);
  });
  const sorters = { added: (a, b) => b.addedAt - a.addedAt, applied: (a, b) => (b.appliedAt || 0) - (a.appliedAt || 0), match: (a, b) => (b.match?.score ?? -1) - (a.match?.score ?? -1), company: (a, b) => a.company.localeCompare(b.company), status: (a, b) => STATUSES.findIndex((x) => x[0] === a.status) - STATUSES.findIndex((x) => x[0] === b.status) };
  rows.sort(sorters[T.sort] || sorters.added);
  const filters = [["all", "All"], ["toapply", "To apply"], ["applied", "Applied"], ["active", "In progress"], ["interview", "Interview"], ["offer", "Offer"], ["rejected", "Rejected"], ["dismissed", "Not interested"]];

  return `
  <h1>Tracker</h1>
  <p class="sub">Every job you've picked or applied to. Change the status when you hear back.</p>
  <div class="kpis">
    ${kpi("Jobs tracked", st.live.length, `${st.toApply.length} still to apply`)}
    ${kpi("Applied", st.applied.length, `${weeks[7].n} this week`)}
    ${kpi("Waiting to hear", st.waiting.length, `${st.waiting.filter((j) => daysSince(j.appliedAt) >= 14).length} over 2 weeks`)}
    ${kpi("Heard back", st.heard.length, rate == null ? "no applications yet" : `${rate}% response rate`)}
    ${kpi("Interviews", st.interviews.length, st.applied.length ? `${Math.round((st.interviews.length / st.applied.length) * 100)}% of applications` : "—")}
    ${kpi("Offers", st.offers.length, st.offers.length ? "🎉 nice work" : "keep going")}
  </div>
  <div class="charts">
    <div class="card"><h2>Applications per week</h2><p class="hint">Last 8 weeks, by date applied</p>
      <div class="vbars">${weeks.map((w) => `<div class="c" data-tip="Week of ${esc(fmtDate(w.start))}: ${plural(w.n, "application")}"><div class="v">${w.n || ""}</div><div class="b" style="height:${(w.n / wmax) * 82}%"></div></div>`).join("")}</div>
      <div class="vlabels">${weeks.map((w, i) => `<span>${i === 7 ? "This wk" : new Date(w.start).toLocaleDateString(undefined, { day: "numeric", month: "short" })}</span>`).join("")}</div>
    </div>
    <div class="card"><h2>Where your jobs are</h2><p class="hint">Current status of every tracked job</p>
      <div class="hbars">${scounts.map((x) => `<div class="hb" data-tip="${esc(x.l)}: ${plural(x.n, "job")}"><span>${esc(x.l)}</span><div class="track"><div class="fill" style="width:${(x.n / smax) * 100}%"></div></div><span class="n">${x.n}</span></div>`).join("")}</div>
    </div>
  </div>
  <div class="toolbar">
    <div class="chips">${filters.map(([v, l]) => `<button class="chip ${T.filter === v ? "on" : ""}" data-act="tFilter" data-v="${v}">${l}</button>`).join("")}</div>
    <div class="sp"></div>
    <input type="search" id="tq" placeholder="Search jobs, companies, notes" value="${esc(T.q)}">
    <select id="tsort" data-change="tsort">${[["added", "Recently added"], ["applied", "Recently applied"], ["match", "Best match"], ["status", "Status"], ["company", "Company A–Z"]].map(([v, l]) => `<option value="${v}" ${T.sort === v ? "selected" : ""}>${l}</option>`).join("")}</select>
    <button class="btn sm" data-act="manualAdd" data-status="applied">+ Add job</button>
    <button class="btn sm" data-act="csv">Export CSV</button>
  </div>
  ${rows.length ? `<div class="tablewrap"><table>
    <thead><tr><th>Job</th><th>Location</th><th>Match</th><th>Added</th><th>Applied</th><th>Status</th><th>Documents</th><th style="width:22%">Notes</th><th></th></tr></thead>
    <tbody>${rows.map(trackerRow).join("")}</tbody></table></div>` :
    `<div class="card empty"><b>${st.all.length ? "No jobs match this filter" : "No jobs tracked yet"}</b>${st.all.length ? "Try another filter or clear the search." : `Pick jobs on <a href="#find">Find jobs</a>, or <a href="#" data-act="manualAdd" data-status="applied">add one you've already applied to</a>.`}</div>`}`;
}
function trackerRow(j) {
  const d = daysSince(j.appliedAt);
  const follow = j.status === "applied" && d >= 14;
  return `<tr>
    <td>${j.url ? `<a class="jt" href="${esc(j.url)}" target="_blank" rel="noopener">${esc(j.title)}</a>` : `<span class="jt">${esc(j.title)}</span>`}<div class="muted small">${esc(j.company)}${j.source === "manual" ? " · added by you" : ""}</div></td>
    <td class="small">${esc(j.location || "")}</td>
    <td>${j.match?.score != null ? `<span class="pill ${j.match.score >= 75 ? "good" : j.match.score >= 55 ? "acc" : ""}">${esc(j.match.score)}</span>` : `<span class="muted">—</span>`}</td>
    <td class="small">${fmtDate(j.addedAt)}</td>
    <td class="small">${APPLIED_SET.has(j.status) ? `<input type="date" value="${isoDate(j.appliedAt)}" data-change="appliedDate" data-key="${esc(j.key)}" aria-label="Date applied"><div class="muted">${d != null ? ago(j.appliedAt) : ""}</div>` :
      TO_APPLY.has(j.status) ? `<button class="btn sm good" data-act="markApplied" data-key="${esc(j.key)}">Mark applied</button>` : `<span class="muted">—</span>`}</td>
    <td><select class="st" data-change="status" data-key="${esc(j.key)}" aria-label="Status">${STATUSES.map(([v, l]) => `<option value="${v}" ${j.status === v ? "selected" : ""}>${l}</option>`).join("")}</select>
      ${follow ? `<div><span class="pill warn" data-tip="No reply after 2 weeks. A short follow-up email can help.">Follow up?</span></div>` : ""}</td>
    <td>${j.tailored ? `<button class="btn sm ghost" data-act="dl" data-key="${esc(j.key)}" data-kind="resume" data-fmt="docx">Resume</button><button class="btn sm ghost" data-act="dl" data-key="${esc(j.key)}" data-kind="letter" data-fmt="docx">Letter</button>`
      : TO_APPLY.has(j.status) ? `<button class="btn sm" data-act="openWs" data-key="${esc(j.key)}">Tailor</button>` : `<span class="muted small">—</span>`}</td>
    <td><input class="note" type="text" value="${esc(j.notes || "")}" placeholder="Add a note" data-change="note" data-key="${esc(j.key)}"></td>
    <td><button class="btn sm ghost" data-act="delJob" data-key="${esc(j.key)}" aria-label="Delete" data-tip="Remove from tracker">✕</button></td>
  </tr>`;
}
function exportCsv() {
  const cols = ["Title", "Company", "Location", "Status", "Match", "Added", "Applied", "Responded", "Link", "Notes"];
  const q = (v) => `"${String(v == null ? "" : v).replace(/"/g, '""')}"`;
  const rows = [...S.jobs.values()].sort((a, b) => b.addedAt - a.addedAt).map((j) => [j.title, j.company, j.location, SLABEL[j.status], j.match?.score ?? "", isoDate(j.addedAt), isoDate(j.appliedAt), isoDate(j.respondedAt), j.url, j.notes].map(q).join(","));
  return [cols.map(q).join(","), ...rows].join("\n");
}

/* ============ actions ============ */
const ACT = {
  saveProfile: async () => {
    const wasDirty = S.profileDirty;
    S.profileMsg = "Saving…"; S.profileDirty = false; refreshSaveBar();
    try { await saveProfile(); S.profileMsg = "Saved just now"; toast(wasDirty ? "Profile saved" : "Profile is saved"); }
    catch { S.profileDirty = true; S.profileMsg = null; toast("Couldn't save. Try again."); }
    refreshSaveBar();
  },
  previewResume: (b) => {
    const r = S.resumes.get(b.dataset.id); if (!r) return;
    const m = $("#modal");
    m.innerHTML = `<div class="mbox" style="max-width:760px"><h3>${esc(r.name)}</h3><p class="small muted">This is the text the tracker read from your file. If something's missing or garbled, paste the text instead.</p>
      <pre style="white-space:pre-wrap;font:12.5px/1.5 inherit;background:var(--surface-2);padding:14px;border-radius:10px;max-height:60vh;overflow:auto">${esc(r.text)}</pre>
      <div class="mfoot"><button class="btn" data-act="closeModal">Close</button></div></div>`;
    m.hidden = false;
  },
  delResume: (b) => {
    const r = S.resumes.get(b.dataset.id); if (!r) return;
    ask(`<h3>Delete “${esc(r.name)}”?</h3><p class="muted" style="margin:0">Documents you already generated stay as they are.</p>`, [{ label: "Cancel" }, { label: "Delete", cls: "pri", fn: () => { deleteResume(r.id); S.profileMsg = `Resume “${r.name}” deleted`; } }]);
  },
  pasteResume: () => {
    const text = tidyText($("#pasteText").value), name = $("#pasteName").value.trim() || "Pasted resume";
    if (text.length < 200) { toast("That looks too short for a resume. Paste the full text."); return; }
    saveResume({ id: "r-" + now().toString(36), name, text, addedAt: now() }); S.profileMsg = `Resume “${name}” added and saved`; toast("Added and saved " + name);
  },
  suggestRoles: async (b) => {
    const msg = $("#rolemsg"); b.disabled = true; if (msg) msg.textContent = "Reading your resume…";
    try {
      const out = await S.sample.json(`Suggest 4 to 6 job titles this person should search for on job boards, based on their resume. Use common job-board titles at the right seniority. Reply with only JSON: {"roles":["Title 1","Title 2"]}\n\nRESUME AND NOTES:\n${candidateMaterial("all", 6000).slice(0, 20000)}`, { modelTier: "quick" });
      const roles = (out && Array.isArray(out.roles) ? out.roles : []).map(String).filter(Boolean).slice(0, 6);
      if (!roles.length) throw { code: "invalid_json" };
      S.profile.prefs.roles = roles.join(", "); S.profileDirty = true; await saveProfile(); S.profileMsg = "Suggested job titles added and saved"; renderView(); toast("Added suggested job titles");
    } catch (e) { if (msg) msg.textContent = sampleErr(e); b.disabled = false; }
  },
  find: () => runFind(),
  selAll: () => { S.find.selected = new Set(S.find.results.map((_, i) => i)); renderView(); },
  selNone: () => { S.find.selected = new Set(); renderView(); },
  hideResult: (b) => {
    const i = +b.dataset.i, r = S.find.results[i]; if (!r) return;
    saveJob({ key: keyFor(r.title, r.company), title: r.title, company: r.company, location: r.location, url: safeUrl(r.url), source: "adzuna", status: "dismissed", addedAt: now(), history: [{ s: "dismissed", at: now() }], notes: "", match: r.score != null ? { score: r.score, reason: r.reason || "" } : null });
    S.find.results = S.find.results.filter((_, k) => k !== i);
    S.find.selected = new Set([...S.find.selected].filter((k) => k !== i).map((k) => (k > i ? k - 1 : k)));
    renderView();
  },
  addSelected: () => addSelected(),
  openWs: (b) => openWs(b.dataset.key),
  closeWs: () => closeWs(),
  closeModal: () => { if (S.ws) closeWs(); else { $("#modal").hidden = true; $("#modal").innerHTML = ""; } },
  wsGenerate: () => wsRun("new"),
  wsRevise: () => wsRun("revise"),
  wsStop: () => S.ws && S.ws.ctl && S.ws.ctl.abort(),
  wsTab: (b) => { S.ws.tab = b.dataset.tab; renderWs(); },
  dl: (b) => downloadDoc(b.dataset.key || (S.ws && S.ws.key), b.dataset.kind, b.dataset.fmt),
  copyDoc: async (b) => {
    const j = S.jobs.get(S.ws.key); if (!j || !j.tailored) return;
    const text = b.dataset.kind === "resume" ? JTDocs.resumeText(j.tailored.resume) : JTDocs.letterText(j.tailored.coverLetter, j.tailored.resume);
    try { await navigator.clipboard.writeText(text); toast("Copied to clipboard"); } catch { toast("Couldn't copy. Use the download instead."); }
  },
  opened: (b) => {
    const key = b.dataset.key, j = S.jobs.get(key); if (!j) return;
    S.pendingAsk = { key, at: now() };
    const nj = clone(j); nj.openedAt = now(); saveJob(nj);
    return "follow"; // let the link open
  },
  markApplied: (b) => {
    const key = b.dataset.key; setStatus(key, "applied");
    if (S.ws && S.ws.key === key) closeWs();
    toast("Marked as applied");
  },
  notYet: (b) => clearOpened(b.dataset.key),
  dismiss: (b) => setStatus(b.dataset.key, "dismissed"),
  manualAdd: (b, e) => { e.preventDefault(); manualAdd(b.dataset.status || "shortlisted"); },
  tFilter: (b) => { S.tracker.filter = b.dataset.v; renderView(); },
  delJob: (b) => {
    const j = S.jobs.get(b.dataset.key); if (!j) return;
    ask(`<h3>Remove this job?</h3><p class="muted" style="margin:0">${esc(j.title)} at ${esc(j.company)}. It may show up again in future searches. To hide it for good, set its status to “Not interested” instead.</p>`,
      [{ label: "Cancel" }, { label: "Set to Not interested", fn: () => setStatus(j.key, "dismissed") }, { label: "Remove", cls: "pri", fn: () => deleteJob(j.key) }]);
  },
  csv: async () => {
    if (!S.downloads) { toast("Downloads aren't available in this view."); return; }
    try { await S.downloads.save({ filename: "job-applications.csv", data: exportCsv() }); } catch (e) { if (e.code !== "declined") toast("Couldn't export."); }
  },
};
const CHANGE = {
  upload: (el) => { const files = [...el.files]; el.value = ""; addResumeFiles(files); },
  count: (el) => { S.find.count = +el.value; },
  pick: (el) => { const i = +el.dataset.i; if (el.checked) S.find.selected.add(i); else S.find.selected.delete(i); renderView(); },
  status: (el) => { setStatus(el.dataset.key, el.value); el.blur(); },
  appliedDate: (el) => { const j = S.jobs.get(el.dataset.key); if (!j || !el.value) return; const nj = clone(j); nj.appliedAt = new Date(el.value + "T12:00:00").getTime(); saveJob(nj); el.blur(); },
  note: (el) => { const j = S.jobs.get(el.dataset.key); if (!j || (j.notes || "") === el.value) return; const nj = clone(j); nj.notes = el.value.slice(0, 1000); saveJob(nj); },
  tsort: (el) => { S.tracker.sort = el.value; el.blur(); renderView(); },
};

document.addEventListener("click", (e) => {
  const b = e.target.closest("[data-act]"); if (!b) return;
  const fn = ACT[b.dataset.act]; if (!fn) return;
  if (b.tagName === "A" && b.dataset.act !== "opened" && b.getAttribute("href") === "#") e.preventDefault();
  if (b.dataset.act === "closeWs" && b.tagName === "A") { closeWs(); return; }
  fn(b, e);
});
document.addEventListener("input", (e) => {
  const t = e.target;
  if (t.id === "wsDesc" && S.ws) { S.ws.desc = t.value; S.ws.descDirty = true; }
  else if (t.id === "wsExtra" && S.ws) S.ws.extra = t.value;
  else if (t.id === "wsRevise" && S.ws) S.ws.revise = t.value;
  else if (t.id === "tq") { S.tracker.q = t.value; clearTimeout(ACT._tq); ACT._tq = setTimeout(() => { const pos = t.selectionStart; renderView(); const n = $("#tq"); if (n) { n.focus(); n.setSelectionRange(pos, pos); } }, 250); }
});
document.addEventListener("change", (e) => {
  if (e.target.id === "wsBase" && S.ws) S.ws.base = e.target.value;
  if (e.target.id === "wsDesc" && S.ws && S.ws.descDirty) {
    const j = S.jobs.get(S.ws.key); if (j) { const nj = clone(j); nj.description = S.ws.desc.slice(0, 15000); nj.descIsSnippet = false; saveJob(nj); S.ws.descDirty = false; }
  }
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") { if (!$("#ask").hidden) { $("#ask").hidden = true; $("#ask").innerHTML = ""; } else if (!$("#modal").hidden) ACT.closeModal(); }
  if (e.key === "Enter" && e.target.id === "wsRevise") wsRun("revise");
});
$("#modal").addEventListener("mousedown", (e) => { if (e.target.id === "modal" && !(S.ws && S.ws.busy)) ACT.closeModal(); });

/* tooltips */
document.addEventListener("mouseover", (e) => {
  const el = e.target.closest && e.target.closest("[data-tip]"), tip = $("#tip");
  if (!el) { tip.hidden = true; return; }
  tip.textContent = el.dataset.tip; tip.hidden = false;
  const r = el.getBoundingClientRect(), w = tip.offsetWidth;
  tip.style.left = Math.max(8, Math.min(window.innerWidth - w - 8, r.left + r.width / 2 - w / 2)) + "px";
  tip.style.top = Math.max(8, r.top - tip.offsetHeight - 8) + "px";
});
document.addEventListener("scroll", () => ($("#tip").hidden = true), true);

boot();
})();
