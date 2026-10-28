/*
 * Antwort-Übersicht (Web) – liest ein OneNote-Kursnotizbuch über Microsoft Graph
 * und zeigt die Antworten aller Schüler:innen zu einer verteilten Seite an.
 *
 * - Anmeldung: Microsoft-Konto der Schule (OAuth 2.0 mit PKCE, ohne Zusatzbibliothek)
 * - Rechte: nur LESEN (Notes.Read.All, Team.ReadBasic.All, User.Read)
 * - Alle Daten bleiben in deinem Browser; es gibt keinen eigenen Server.
 */
"use strict";

const CFG = window.APP_CONFIG || {};
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = "openid profile offline_access User.Read Team.ReadBasic.All " +
  (CFG.feedbackSchreiben ? "Notes.ReadWrite.All" : "Notes.Read.All");
const AUTH = `https://login.microsoftonline.com/${CFG.tenantId}/oauth2/v2.0`;
const REDIRECT = location.origin + location.pathname;
const SKIP_PREFIX = "_";

const $ = (id) => document.getElementById(id);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const LS = {
  get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* egal */ } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* egal */ } },
};

const state = {
  base: "",              // "/me/onenote" oder "/groups/{id}/onenote"
  students: [],          // [{ name, sectionIds: [] }]
  sectionToStudent: new Map(),
  results: [],
  title: "",
  loading: false,
  contentCache: new Map(),  // Seiten-ID -> { modified, html }
  knownPages: new Map(),    // Seitentitel -> Map(Schüler -> Seite)
  noSearch: new Set(),      // Orte, an denen die Sammelsuche nicht funktioniert
};

/* ------------------------------------------------------------------ */
/* Anmeldung                                                           */
/* ------------------------------------------------------------------ */

class NeedLogin extends Error {}

function b64url(bytes) {
  let s = "";
  bytes.forEach((b) => { s += String.fromCharCode(b); });
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function randomString(n) {
  const a = new Uint8Array(n);
  crypto.getRandomValues(a);
  return b64url(a);
}

async function login() {
  const verifier = randomString(48);
  const st = randomString(16);
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  LS.set("aue.pkce", JSON.stringify({ verifier, state: st }));
  const p = new URLSearchParams({
    client_id: CFG.clientId,
    response_type: "code",
    redirect_uri: REDIRECT,
    response_mode: "query",
    scope: SCOPES,
    state: st,
    code_challenge: b64url(new Uint8Array(hash)),
    code_challenge_method: "S256",
    prompt: "select_account",
  });
  location.assign(`${AUTH}/authorize?${p}`);
}

function logout() {
  LS.del("aue.token");
  location.assign(`${AUTH}/logout?post_logout_redirect_uri=${encodeURIComponent(REDIRECT)}`);
}

async function handleRedirect() {
  const q = new URLSearchParams(location.search);
  if (q.has("error")) {
    history.replaceState(null, "", REDIRECT);
    const d = (q.get("error_description") || q.get("error")).split(/\r?\n/)[0];
    throw new Error("Anmeldung fehlgeschlagen: " + d);
  }
  if (!q.has("code")) return;
  const saved = JSON.parse(LS.get("aue.pkce") || "null");
  LS.del("aue.pkce");
  history.replaceState(null, "", REDIRECT);
  if (!saved || saved.state !== q.get("state")) {
    throw new Error("Anmeldung konnte nicht abgeschlossen werden. Bitte noch einmal anmelden.");
  }
  await tokenRequest({
    grant_type: "authorization_code",
    code: q.get("code"),
    redirect_uri: REDIRECT,
    code_verifier: saved.verifier,
  });
}

async function tokenRequest(params) {
  const r = await fetch(`${AUTH}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CFG.clientId, scope: SCOPES, ...params }),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) {
    const d = (j.error_description || j.error || `HTTP ${r.status}`).split(/\r?\n/)[0];
    throw new Error(d);
  }
  const tok = {
    access: j.access_token,
    refresh: j.refresh_token || params.refresh_token || null,
    exp: Date.now() + (Number(j.expires_in || 3600) - 120) * 1000,
  };
  LS.set("aue.token", JSON.stringify(tok));
  return tok;
}

async function getAccessToken() {
  const t = JSON.parse(LS.get("aue.token") || "null");
  if (!t) return null;
  if (Date.now() < t.exp) return t.access;
  if (t.refresh) {
    try { return (await tokenRequest({ grant_type: "refresh_token", refresh_token: t.refresh })).access; }
    catch (e) { console.warn("Token-Erneuerung fehlgeschlagen", e); }
  }
  LS.del("aue.token");
  return null;
}

function expireToken() {
  const t = JSON.parse(LS.get("aue.token") || "null");
  if (t) { t.exp = 0; LS.set("aue.token", JSON.stringify(t)); }
}

/* ------------------------------------------------------------------ */
/* Microsoft Graph                                                     */
/* ------------------------------------------------------------------ */

// Wird Microsoft zu viel, pausieren ALLE Anfragen gemeinsam (statt dass jede einzeln
// weiterdrängelt und die Sperre verlängert).
let graphPauseUntil = 0;
let graphRequests = 0;

async function graph(path, mode, init) {
  const url = path.startsWith("http") ? path : GRAPH + path;
  for (let attempt = 0; ; attempt++) {
    while (Date.now() < graphPauseUntil) await sleep(graphPauseUntil - Date.now());
    const token = await getAccessToken();
    if (!token) throw new NeedLogin();
    const opt = init || {};
    graphRequests++;
    const r = await fetch(url, {
      method: opt.method || "GET",
      headers: { Authorization: "Bearer " + token, ...(opt.headers || {}) },
      body: opt.body,
    });
    if (r.status === 401 && attempt === 0) { expireToken(); continue; }
    if ([429, 503, 504].includes(r.status) && attempt < 6) {
      const wait = Number(r.headers.get("Retry-After")) || Math.min(60, 10 * 2 ** attempt);
      graphPauseUntil = Math.max(graphPauseUntil, Date.now() + wait * 1000);
      setStatus(`Microsoft begrenzt die Anzahl der Abfragen – kurze Pause (${wait} s) …`);
      continue;
    }
    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try { const j = await r.json(); msg = (j.error && j.error.message) || msg; } catch (e) { /* egal */ }
      const err = new Error(msg);
      err.status = r.status;
      throw err;
    }
    if (mode === "none") return null;
    return mode === "text" ? r.text() : mode === "blob" ? r.blob() : r.json();
  }
}

async function graphAll(path) {
  const out = [];
  let next = path;
  while (next) {
    const j = await graph(next);
    out.push(...(j.value || []));
    next = j["@odata.nextLink"];
  }
  return out;
}

async function pool(items, n, fn) {
  const res = new Array(items.length);
  let i = 0;
  const worker = async () => { while (i < items.length) { const k = i++; res[k] = await fn(items[k], k); } };
  await Promise.all(Array.from({ length: Math.min(n, items.length) }, worker));
  return res;
}

/* ------------------------------------------------------------------ */
/* Oberfläche: Auswahl                                                 */
/* ------------------------------------------------------------------ */

function setStatus(msg, isError) {
  $("status").textContent = msg || "";
  $("status").classList.toggle("error", !!isError);
}

function fillSelect(sel, items, savedValue) {
  sel.innerHTML = items.map((it) => `<option value="${esc(it.value)}">${esc(it.label)}</option>`).join("");
  if (savedValue && items.some((it) => it.value === savedValue)) sel.value = savedValue;
  sel.disabled = !items.length;
}

async function afterLogin() {
  $("loginView").hidden = true;
  $("setupView").hidden = false;

  const me = await graph("/me?$select=displayName");
  $("user").innerHTML = `<span>${esc(me.displayName || "")}</span><button id="logoutBtn" type="button">Abmelden</button>`;
  $("logoutBtn").addEventListener("click", logout);

  const locations = [{ value: "/me/onenote", label: "Mein OneDrive" }];
  try {
    const teams = await graphAll("/me/joinedTeams?$select=id,displayName");
    teams.sort((a, b) => a.displayName.localeCompare(b.displayName, "de"))
      .forEach((t) => locations.push({ value: `/groups/${t.id}/onenote`, label: "Team: " + t.displayName }));
  } catch (e) {
    console.warn("Teams konnten nicht geladen werden", e);
  }
  fillSelect($("locationSel"), locations, LS.get("aue.loc"));
  await onLocation();
}

async function onLocation() {
  state.base = $("locationSel").value;
  LS.set("aue.loc", state.base);
  $("notebookSel").disabled = true;
  setTitleEnabled(false);
  setStatus("Lade Notizbücher …");
  const nbs = await graphAll(`${state.base}/notebooks?$select=id,displayName`);
  nbs.sort((a, b) => a.displayName.localeCompare(b.displayName, "de"));
  fillSelect($("notebookSel"), nbs.map((n) => ({ value: n.id, label: n.displayName })), LS.get("aue.nb"));
  if (!nbs.length) { setStatus("Hier wurden keine Notizbücher gefunden."); return; }
  await onNotebook();
}

async function onNotebook() {
  const nbId = $("notebookSel").value;
  LS.set("aue.nb", nbId);
  setTitleEnabled(false);
  setStatus("Lese Schülerliste …");

  const groups = await graphAll(
    `${state.base}/notebooks/${nbId}/sectionGroups?$expand=sections,sectionGroups($expand=sections)`
  );
  state.students = groups
    .filter((g) => !(g.displayName || "").startsWith(SKIP_PREFIX))
    .map((g) => {
      const ids = (g.sections || []).map((s) => s.id);
      (g.sectionGroups || []).forEach((sg) => (sg.sections || []).forEach((s) => ids.push(s.id)));
      return { name: g.displayName, sectionIds: ids };
    })
    .sort((a, b) => a.name.localeCompare(b.name, "de"));
  state.sectionToStudent = new Map();
  state.students.forEach((s) => s.sectionIds.forEach((id) => state.sectionToStudent.set(id, s)));

  if (!state.students.length) {
    setStatus("Keine Schüler-Abschnittsgruppen gefunden. Ist das ein Kursnotizbuch?", true);
    return;
  }

  // Seitentitel-Vorschläge aus den Abschnitten der ersten Person
  setStatus("Lade Seitentitel …");
  const first = state.students[0];
  const lists = await pool(first.sectionIds, 4, (id) =>
    graphAll(`${state.base}/sections/${id}/pages?$select=title&$top=100`).catch(() => [])
  );
  const titles = [...new Set(lists.flat().map((p) => (p.title || "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "de"));
  $("titleList").innerHTML = titles.map((t) => `<option value="${esc(t)}"></option>`).join("");

  setTitleEnabled(true);
  setStatus(`${state.students.length} Schüler:innen gefunden. Seitentitel wählen und „Antworten laden“ tippen.`);
}

function setTitleEnabled(on) {
  $("titleInput").disabled = !on;
  $("loadBtn").disabled = !on;
}

/* ------------------------------------------------------------------ */
/* Antworten laden                                                     */
/* ------------------------------------------------------------------ */

function norm(s) {
  return (s || "").replace(/\s+/g, " ").trim().toLowerCase();
}

async function findPages(title, quick) {
  const wanted = norm(title);
  const found = new Map(); // student -> page
  const take = (student, page) => {
    if (!student || norm(page.title) !== wanted) return;
    const prev = found.get(student);
    if (!prev || (page.lastModifiedDateTime || "") > (prev.lastModifiedDateTime || "")) found.set(student, page);
  };

  // 1) schnell: eine Suche über alle Seiten (wird übersprungen, wenn sie hier schon einmal scheiterte)
  if (!state.noSearch.has(state.base)) try {
    const filter = `contains(tolower(title),'${title.toLowerCase().replace(/'/g, "''")}')`;
    const pages = await graphAll(
      `${state.base}/pages?$filter=${encodeURIComponent(filter)}` +
      `&$select=id,title,links,lastModifiedDateTime&$expand=parentSection&$top=100`
    );
    pages.forEach((p) => take(state.sectionToStudent.get(p.parentSection && p.parentSection.id), p));
  } catch (e) {
    if (e instanceof NeedLogin) throw e;
    console.warn("Sammelsuche nicht möglich – suche pro Abschnitt", e);
    state.noSearch.add(state.base);
  }
  if (found.size) return found;

  // 2) Bereits gefundene Seiten nur kurz auf Änderungen prüfen (1 kleine Abfrage pro Seite)
  const known = state.knownPages.get(wanted) || new Map();
  await pool([...known], 2, async ([s, p]) => {
    try {
      take(s, await graph(`${state.base}/pages/${p.id}?$select=id,title,links,lastModifiedDateTime`));
    } catch (e) {
      if (e instanceof NeedLogin) throw e;   // Seite gelöscht o. Ä. → unten neu suchen
    }
  });

  // 3) Abschnitte durchsuchen – nur bei Personen, deren Seite noch nicht bekannt ist,
  //    und nicht beim automatischen Aktualisieren
  if (!quick || !known.size) {
    const jobs = [];
    state.students.forEach((s) => { if (!found.has(s)) s.sectionIds.forEach((id) => jobs.push({ s, id })); });
    let done = 0;
    await pool(jobs, 2, async ({ s, id }) => {
      const pages = await graphAll(
        `${state.base}/sections/${id}/pages?$select=id,title,links,lastModifiedDateTime&$top=100`
      ).catch((e) => { if (e instanceof NeedLogin) throw e; return []; });
      pages.forEach((p) => take(s, p));
      setStatus(`Durchsuche Abschnitte … ${++done}/${jobs.length}`);
    });
  }
  state.knownPages.set(wanted, new Map(found));
  return found;
}

async function loadAnswers(auto) {
  if (state.loading) return;
  const title = $("titleInput").value.trim();
  const markerRaw = $("markerInput").value.trim();
  if (!title) { setStatus("Bitte einen Seitentitel eingeben.", true); return; }
  if (!markerRaw) { setStatus("Bitte die erste Zeile des Antwortfelds angeben.", true); return; }
  LS.set("aue.title", title);
  LS.set("aue.marker", markerRaw);

  state.loading = true;
  $("loadBtn").disabled = true;
  setStatus(`Suche „${title}“ …`);
  try {
    const before = graphRequests;
    const found = await findPages(title, auto === true);
    const markerBase = norm(markerRaw).replace(/[:\s]+$/, "");
    const changed = [...found.values()].filter((p) => {
      const c = state.contentCache.get(p.id);
      return !c || c.modified !== p.lastModifiedDateTime;
    }).length;
    setStatus(changed ? `Lese ${changed} geänderte Seiten …` : "Keine Änderungen – prüfe …");
    state.results = await pool(state.students, 2, async (s) => {
      const page = found.get(s);
      if (!page) return { name: s.name, status: "keine-seite", fields: [] };
      let cached = state.contentCache.get(page.id);
      if (!cached || cached.modified !== page.lastModifiedDateTime || !page.lastModifiedDateTime) {
        const fresh = await graph(`${state.base}/pages/${page.id}/content?includeIDs=true`, "text");
        cached = { modified: page.lastModifiedDateTime, html: fresh };
        state.contentCache.set(page.id, cached);
      }
      const html = cached.html;
      const { fields, taskText } = evaluateHtml(html, markerBase);
      for (const f of fields) {
        f.images = (await Promise.all(f.imageSrcs.slice(0, 4).map(loadImage))).filter(Boolean);
      }
      return { name: s.name, page, fields, taskText, status: statusOf(fields) };
    });
    state.title = title;
    buildFieldOptions();
    const time = new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
    const n = graphRequests - before;
    setStatus(`Stand: ${time} Uhr · ${n} ${n === 1 ? "Abfrage" : "Abfragen"} an Microsoft`);
    render();
  } catch (e) {
    handleError(e);
  } finally {
    state.loading = false;
    $("loadBtn").disabled = false;
  }
}

/* ------------------------------------------------------------------ */
/* Seiteninhalt auswerten                                              */
/* ------------------------------------------------------------------ */

function clean(s) { return (s || "").replace(/\s+/g, " ").trim(); }

// Liest die Textzeilen aus einem Textfeld (Container) der Seite
function collect(root) {
  const lines = [];
  const extras = new Set();
  const images = [];
  const addImg = (img) => {
    const src = img.getAttribute("src") || img.getAttribute("data-fullres-src");
    if (src) images.push(src);
  };
  const walk = (el) => {
    for (const n of el.childNodes) {
      if (n.nodeType === 3) { const t = clean(n.textContent); if (t) lines.push(t); continue; }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (/^(P|H[1-6])$/.test(tag)) {
        const t = clean(n.textContent);
        if (t) lines.push(t);
        n.querySelectorAll("img").forEach(addImg);
      } else if (tag === "LI") {
        const c = n.cloneNode(true);
        c.querySelectorAll("ul,ol").forEach((x) => x.remove());
        const t = clean(c.textContent);
        if (t) lines.push("• " + t);
        n.querySelectorAll(":scope > ul, :scope > ol").forEach(walk);
      } else if (tag === "TABLE") {
        extras.add("Tabelle");
        n.querySelectorAll("tr").forEach((tr) => {
          const cells = [...tr.children].map((td) => clean(td.textContent)).filter(Boolean);
          if (cells.length) lines.push(cells.join(" | "));
        });
      } else if (tag === "IMG") {
        addImg(n);
      } else if (tag === "OBJECT" || tag === "IFRAME") {
        extras.add("Anhang");
      } else if (tag !== "BR" && tag !== "SCRIPT" && tag !== "STYLE") {
        walk(n);
      }
    }
  };
  walk(root);
  return { lines, extras: [...extras], images };
}

// Prüft, ob eine Zeile ein Antwortfeld einleitet: "Answer:", "Answer a:", "Answer 2: Text" …
function matchMarker(line, base) {
  if (!norm(line).startsWith(base)) return null;
  const rest = clean(line).slice(base.length);
  const colon = rest.indexOf(":");
  if (colon >= 0 && colon <= 12) {
    return { label: rest.slice(0, colon).trim(), answer: rest.slice(colon + 1).trim() };
  }
  if (/^[\p{L}\p{N}]/u.test(rest)) return null; // z. B. "Answering …" ist kein Feld
  return { label: "", answer: rest.trim() };
}

// Position eines Textfelds auf der Seite (für die Reihenfolge von oben nach unten)
function posOf(el) {
  const st = el.getAttribute("style") || "";
  const get = (k) => {
    const m = new RegExp(`(?:^|;)\\s*${k}\\s*:\\s*(-?[\\d.]+)`).exec(st);
    return m ? parseFloat(m[1]) : 0;
  };
  return { top: get("top"), left: get("left") };
}

// Liefert alle Antwortfelder der Seite, sortiert von oben nach unten
function evaluateHtml(html, markerBase) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  let boxes = [...doc.body.children].filter((e) => e.tagName === "DIV");
  if (!boxes.length) boxes = [doc.body];
  boxes = boxes
    .map((el, i) => ({ el, i, ...posOf(el) }))
    .sort((a, b) => a.top - b.top || a.left - b.left || a.i - b.i)
    .map((b) => b.el);

  const fields = [];
  const taskLines = [];
  const fbBoxes = [];   // Textfelder auf der Vorlage, die mit "Feedback:" beginnen
  for (const box of boxes) {
    const { lines, extras, images } = collect(box);
    if (!lines.length) continue;
    const m = matchMarker(lines[0], markerBase);
    if (!m) {
      const fb = markerBase !== "feedback" && matchMarker(lines[0], "feedback");
      if (fb) {
        fbBoxes.push({
          label: fb.label, id: box.getAttribute("id") || "",
          lines: (fb.answer ? [fb.answer] : []).concat(lines.slice(1)),
        });
      } else {
        taskLines.push(...lines);
      }
      continue;
    }
    // Bereits eingetragenes Feedback gehört nicht zur Antwort
    let body = lines.slice(1);
    let existingFeedback = "";
    const fi = body.findIndex(isFeedbackLine);
    if (fi >= 0) {
      existingFeedback = body.slice(fi).join("\n").replace(/^💬\s*feedback:\s*/i, "");
      body = body.slice(0, fi);
    }
    const text = (m.answer ? [m.answer] : []).concat(body).join("\n");
    fields.push({
      label: m.label, text, extras, imageSrcs: images, existingFeedback,
      targetId: box.getAttribute("id") || "",
      status: text || extras.length || images.length ? "ok" : "leer",
    });
  }
  fields.forEach((f, i) => {
    f.key = f.label ? "L:" + norm(f.label) : "N:" + (i + 1);
    f.display = f.label || String(i + 1);
    f.fbTarget = f.targetId;   // Standard: unter die Antwort
    f.fbBox = null;            // "dedicated" | "shared" wenn ein Feedback-Feld genutzt wird
  });
  assignFeedbackBoxes(fields, fbBoxes);
  return { fields, taskText: taskLines.join("\n").slice(0, 3000) };
}

// Ordnet jedem Antwortfeld ein Feedback-Feld der Vorlage zu:
//  "Feedback a:" ↔ "Answer a:"; sonst der Reihe nach; gibt es nur eins, teilen sich alle Antworten dieses Feld.
function assignFeedbackBoxes(fields, boxes) {
  boxes = boxes.filter((b) => b.id);
  if (!boxes.length || !fields.length) return;
  const used = new Set();
  const open = [];
  fields.forEach((f) => {
    const b = f.label && boxes.find((x) => !used.has(x) && x.label && norm(x.label) === norm(f.label));
    if (b) { used.add(b); useBox(f, b, "dedicated"); } else open.push(f);
  });
  const rest = boxes.filter((b) => !used.has(b));
  if (!open.length) return;
  if (rest.length === open.length) {
    open.forEach((f, i) => useBox(f, rest[i], "dedicated"));
  } else if (boxes.length === 1) {
    open.forEach((f) => useBox(f, boxes[0], fields.length > 1 ? "shared" : "dedicated"));
  }
}

function useBox(f, box, kind) {
  f.fbTarget = box.id;
  f.fbBox = kind;
  const strip = (l) => l.replace(/^💬\s*feedback:\s*/i, "");
  if (kind === "dedicated") {
    f.existingFeedback = box.lines.map(strip).join("\n");
  } else {
    const prefix = norm(f.display) + ")";
    const own = box.lines.filter((l) => norm(l).startsWith(prefix));
    f.existingFeedback = own.map((l) => clean(l).slice(prefix.length).trim()).join("\n");
  }
}

function statusOf(fields) {
  if (!fields.length) return "kein-feld";
  const ok = fields.filter((f) => f.status === "ok").length;
  return ok === fields.length ? "ok" : ok === 0 ? "leer" : "teilweise";
}

// Bilder aus dem Antwortfeld laden (mit Anmeldung) und zwischenspeichern
const imageCache = new Map();
async function loadImage(src) {
  const m = /resources\/([^/?]+)\/\$value/.exec(src);
  if (!m) return null;
  const id = m[1];
  if (!imageCache.has(id)) {
    imageCache.set(id, graph(`${state.base}/resources/${id}/$value`, "blob")
      .then((b) => URL.createObjectURL(b))
      .catch((e) => { console.warn("Bild nicht ladbar", e); imageCache.delete(id); return null; }));
  }
  return imageCache.get(id);
}

/* ------------------------------------------------------------------ */
/* Anzeige                                                             */
/* ------------------------------------------------------------------ */

// Auswahl "Feld": alle Felder, die auf irgendeiner Seite vorkommen
function buildFieldOptions() {
  const keys = new Map();
  state.results.forEach((r) => (r.fields || []).forEach((f) => { if (!keys.has(f.key)) keys.set(f.key, f.display); }));
  const sel = $("fieldSel");
  const prev = sel.value;
  sel.innerHTML = `<option value="">Alle Felder</option>` +
    [...keys].map(([k, d]) => `<option value="${esc(k)}">Feld ${esc(d)}</option>`).join("");
  sel.value = keys.has(prev) ? prev : "";
  $("fieldWrap").hidden = keys.size < 2;
}

// Was für eine Person angezeigt wird – abhängig von der Feldauswahl
function viewOf(r) {
  if (r.status === "keine-seite") return { status: "keine-seite", fields: [] };
  const key = $("fieldWrap").hidden ? "" : $("fieldSel").value;
  const fields = key ? r.fields.filter((f) => f.key === key) : r.fields;
  return { status: statusOf(fields), fields };
}

function badgeText(v) {
  switch (v.status) {
    case "ok": return "beantwortet";
    case "teilweise": return `${v.fields.filter((f) => f.status === "ok").length}/${v.fields.length} beantwortet`;
    case "leer": return "noch nichts";
    case "kein-feld": return "Feld fehlt";
    default: return "Seite fehlt";
  }
}

function render() {
  const res = state.results;
  if (!res.length) { $("results").innerHTML = ""; $("summary").innerHTML = ""; return; }

  const views = res.map((r) => ({ r, v: viewOf(r) }));
  const full = views.filter((x) => x.v.status === "ok").length;
  const part = views.filter((x) => x.v.status === "teilweise").length;
  const pct = Math.round((full / res.length) * 100);
  const selOpt = $("fieldWrap").hidden || !$("fieldSel").value ? "" : " · " + $("fieldSel").selectedOptions[0].text;
  $("summary").innerHTML =
    `${esc(state.title)}${esc(selOpt)} – ${full} von ${res.length} fertig` +
    (part ? `, ${part} teilweise` : "") +
    `<div class="bar"><span style="width:${pct}%"></span></div>`;

  const shown = $("onlyMissing").checked ? views.filter((x) => x.v.status !== "ok") : views;
  $("fbAllBtn").hidden = false;
  fbIndex.clear();
  if (!shown.length) { $("results").innerHTML = `<p class="note">Alle haben geantwortet. 🎉</p>`; return; }

  $("results").innerHTML = shown.map(({ r, v }) => {
    const multi = v.fields.length > 1 || (r.fields && r.fields.length > 1 && v.fields.length === 1);
    const body = v.fields.map((f) => `<div class="field ${f.status}">
        ${multi ? `<div class="flabel">Feld ${esc(f.display)}</div>` : ""}
        ${f.text ? `<p class="answer">${esc(f.text)}</p>` : ""}
        ${(f.images || []).length ? `<div class="thumbs">${f.images.map((u) => `<a href="${esc(u)}" target="_blank" rel="noopener"><img src="${esc(u)}" alt="Bild von ${esc(r.name)}"></a>`).join("")}</div>` : ""}
        ${f.extras.length ? `<div class="note">📎 Enthält: ${esc(f.extras.join(", "))}</div>` : ""}
        ${f.status === "leer" && multi ? `<div class="empty">noch leer</div>` : ""}
        ${feedbackHtml(r, f)}
      </div>`).join("");

    const notes = [];
    if (v.status === "leer") notes.push("Falls mit dem Stift direkt auf die Seite geschrieben: Diese Handschrift wird hier nicht angezeigt – bitte Seite öffnen.");
    if (v.status === "kein-feld") notes.push("Antwortfeld nicht gefunden – evtl. gelöscht oder erste Zeile verändert.");
    if (v.status === "keine-seite") notes.push("Seite bei dieser Person nicht gefunden (umbenannt?).");
    if (r.page && r.page.lastModifiedDateTime) {
      const d = new Date(r.page.lastModifiedDateTime);
      notes.push("Zuletzt geändert: " + d.toLocaleString("de-DE", { weekday: "short", hour: "2-digit", minute: "2-digit" }));
    }
    const l = (r.page && r.page.links) || {};
    const app = l.oneNoteClientUrl && l.oneNoteClientUrl.href;
    const web = l.oneNoteWebUrl && l.oneNoteWebUrl.href;
    const links = [
      app ? `<a href="${esc(app)}">In OneNote öffnen</a>` : "",
      web ? `<a href="${esc(web)}" target="_blank" rel="noopener">Im Web</a>` : "",
    ].join("");

    return `<article class="card ${v.status}">
      <h2><span>${esc(r.name)}</span><span class="badge">${badgeText(v)}</span></h2>
      ${body}
      ${notes.map((n) => `<div class="note">${esc(n)}</div>`).join("")}
      ${links ? `<div class="links">${links}</div>` : ""}
    </article>`;
  }).join("");
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function handleError(e) {
  console.error(e);
  if (e instanceof NeedLogin) {
    $("setupView").hidden = true;
    $("loginView").hidden = false;
    setStatus("Bitte (erneut) anmelden.");
    return;
  }
  let msg = e.message || String(e);
  if (e.status === 403) msg += " – Zugriff verweigert. Wurden die Berechtigungen für die App freigegeben?";
  setStatus("Fehler: " + msg, true);
}

/* ------------------------------------------------------------------ */
/* Start                                                               */
/* ------------------------------------------------------------------ */

async function init() {
  $("loginBtn").addEventListener("click", () => login().catch(handleError));
  $("locationSel").addEventListener("change", () => onLocation().catch(handleError));
  $("notebookSel").addEventListener("change", () => onNotebook().catch(handleError));
  $("loadBtn").addEventListener("click", () => loadAnswers(false));
  $("titleInput").addEventListener("keydown", (e) => { if (e.key === "Enter") loadAnswers(false); });
  $("onlyMissing").addEventListener("change", render);
  $("fieldSel").addEventListener("change", render);
  initAISettings();
  setInterval(() => {
    if ($("autoRefresh").checked && !state.loading && !$("loadBtn").disabled && state.results.length &&
        !(document.activeElement && document.activeElement.tagName === "TEXTAREA")) loadAnswers(true);
  }, 120000);

  $("titleInput").value = LS.get("aue.title") || "";
  $("markerInput").value = LS.get("aue.marker") || "Answer:";

  if (!CFG.clientId || CFG.clientId.startsWith("HIER") || !CFG.tenantId || CFG.tenantId.startsWith("HIER")) {
    setStatus("Bitte zuerst in config.js die Client-ID und Tenant-ID eintragen (siehe README).", true);
    return;
  }

  try {
    await handleRedirect();
  } catch (e) {
    setStatus(e.message, true);
  }
  if (!(await getAccessToken())) {
    $("loginView").hidden = false;
    return;
  }
  try { await afterLogin(); } catch (e) { handleError(e); }
}

init();
