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
  if (!title) { setStatus("Bitte einen Seitentitel eingeben.", true); return; }
  LS.set("aue.title", title);
  LS.set("aue.tasks", $("markerInput").value.trim());

  state.loading = true;
  $("loadBtn").disabled = true;
  setStatus(`Suche „${title}“ …`);
  try {
    const before = graphRequests;
    const found = await findPages(title, auto === true);
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
      return { name: s.name, page, html: cached.html, fields: [] };
    });
    state.title = title;
    await evaluateAll();
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
/* Aufgaben-Auswahl                                                    */
/* ------------------------------------------------------------------ */

// "1) 2), 3a)  B)" → { tokens: {"1","2","3a","B"}, all: false, lower: true }
// Einzelne Buchstaben unterscheiden Groß/klein: "B)" ist Teil B, "b)" eine Unteraufgabe.
function parseTasks(input) {
  const tokens = new Set(
    (input || "").split(/[\s,;]+/)
      .map((t) => t.replace(/^\(/, "").replace(/[).:]+$/, ""))
      .filter((t) => /^(\d{1,2}[a-zA-Z]?|[a-zA-Z])$/.test(t))
      .map((t) => (/^\d/.test(t) ? t.toLowerCase() : t))
  );
  return { tokens, all: tokens.size === 0, lower: [...tokens].some((t) => /[a-z]/.test(t)) };
}

function sameTask(t, x) {
  return /^[a-zA-Z]$/.test(t) ? t === x : t.toLowerCase() === (x || "").toLowerCase();
}

function matchesSelection(f, sel) {
  if (sel.all) return true;
  for (const t of sel.tokens) {
    if (sameTask(t, f.display) || sameTask(t, f.own)) return true;
    // "b)" getippt, auf der Seite gibt es aber nur "B)" (oder umgekehrt) → trotzdem zeigen
    if (/^[a-zA-Z]$/.test(t) && !(state.detectedOwn || new Set()).has(t) &&
        f.own.toLowerCase() === t.toLowerCase() && f.display === f.own) return true;
  }
  return false;
}

// Wertet alle geladenen Seiten (neu) aus. Braucht keine Abfragen an Microsoft –
// deshalb geht das Umschalten der Aufgaben sofort.
async function evaluateAll() {
  const sel = parseTasks($("markerInput").value);
  const detected = new Map();
  state.detectedOwn = new Set();
  const withPage = state.results.filter((r) => r.html);
  withPage.forEach((r) => {
    const p = parsePage(r.html, sel);
    r.allFields = p.fields;
    r.taskText = p.taskText;
    r.fbBoxes = p.fbBoxes;
    p.fields.forEach((f) => { if (!detected.has(f.display)) detected.set(f.display, f); state.detectedOwn.add(f.own); });
  });
  stripTemplate(withPage);

  // "3)" mit Unteraufgaben "3a)", "3b)": Hat bei 3) niemand etwas Eigenes geschrieben,
  // ist es nur die Überschrift → nicht als eigene Aufgabe zählen.
  const keys = new Set(); withPage.forEach((r) => r.allFields.forEach((f) => keys.add(f.key)));
  const parents = [...keys].filter((k) => /^\d/.test(k) && [...keys].some((x) => x !== k && x.startsWith(k) && /^[a-z]/.test(x.slice(k.length))));
  parents.forEach((k) => {
    const all = withPage.flatMap((r) => r.allFields.filter((f) => f.key === k));
    if (all.every((f) => !f.content.length && !f.imageSrcs.length)) {
      withPage.forEach((r) => { r.allFields = r.allFields.filter((f) => f.key !== k); });
    }
  });

  for (const r of withPage) {
    r.fields = r.allFields.filter((f) => matchesSelection(f, sel));
    r.fields.forEach((f) => {
      f.text = f.content.join("\n");
      f.status = f.text || f.imageSrcs.length || f.extras.length ? "ok" : "leer";
    });
    assignFeedbackBoxes(r.fields, r.fbBoxes);
    r.status = statusOf(r.fields);
    for (const f of r.fields) {
      f.images = (await Promise.all(f.imageSrcs.slice(0, 4).map(loadImage))).filter(Boolean);
    }
  }
  state.detected = [...detected.keys()];
  buildTaskChips(sel);
  buildFieldOptions();
}

// Kennzeichen, die auf den Seiten gefunden wurden, als Tipp-Knöpfe
function buildTaskChips(sel) {
  const box = $("taskChips");
  if (!state.detected || !state.detected.length) { box.hidden = true; return; }
  box.hidden = false;
  box.innerHTML = `<span class="note">Auf der Seite gefunden:</span>` + state.detected.map((d) => {
    const on = !sel.all && [...sel.tokens].some((t) => sameTask(t, d));
    return `<button type="button" class="chip${on ? " on" : ""}" data-task="${esc(d)}">${esc(d)})</button>`;
  }).join("") + (sel.all ? "" : `<button type="button" class="chip clear" data-task="">alle</button>`);
}

function toggleTask(d) {
  const input = $("markerInput");
  if (!d) { input.value = ""; }
  else {
    const sel = parseTasks(input.value);
    const hit = [...sel.tokens].find((t) => sameTask(t, d));
    if (hit) sel.tokens.delete(hit); else sel.tokens.add(d);
    // Reihenfolge wie auf der Seite
    const order = state.detected || [];
    const pos = (t) => { const i = order.findIndex((x) => sameTask(t, x)); return i < 0 ? 999 : i; };
    const list = [...sel.tokens].sort((a, b) => pos(a) - pos(b));
    input.value = list.map((t) => (order.find((x) => sameTask(t, x)) || t) + ")").join(" ");
  }
  LS.set("aue.tasks", input.value.trim());
  reevaluate();
}

async function reevaluate() {
  if (!state.results.length || state.loading) return;
  try { await evaluateAll(); render(); } catch (e) { handleError(e); }
}

/* ------------------------------------------------------------------ */
/* Seiteninhalt auswerten                                              */
/* ------------------------------------------------------------------ */

function clean(s) { return (s || "").replace(/\s+/g, " ").trim(); }

function letterOf(n, upper) {
  let s = "";
  for (n = Math.max(1, n); n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(97 + ((n - 1) % 26)) + s;
  return upper ? s.toUpperCase() : s;
}

// Liest ein Textfeld der Seite der Reihe nach aus: Zeilen (mit Element-ID), Bilder, Sonstiges
function collect(root) {
  const items = [];
  const line = (text, el) => { const t = clean(text); if (t) items.push({ kind: "line", text: t, id: (el && el.getAttribute("id")) || "" }); };
  const img = (el) => {
    const src = el.getAttribute("src") || el.getAttribute("data-fullres-src");
    if (src) items.push({ kind: "img", src, id: el.getAttribute("id") || "" });
  };
  const walkList = (list) => {
    const ordered = list.tagName === "OL";
    let n = parseInt(list.getAttribute("start"), 10) || 1;
    for (const li of list.children) {
      if (li.tagName !== "LI") { walk(li); continue; }
      const v = parseInt(li.getAttribute("value"), 10);
      if (v) n = v;
      const style = ((li.getAttribute("style") || "") + ";" + (list.getAttribute("style") || "")).toLowerCase();
      let prefix = "• ";
      if (ordered) {
        const lab = /lower-(alpha|latin)/.test(style) ? letterOf(n, false)
          : /upper-(alpha|latin)/.test(style) ? letterOf(n, true) : String(n);
        prefix = lab + ") ";
      }
      const c = li.cloneNode(true);
      c.querySelectorAll("ul,ol").forEach((x) => x.remove());
      line(prefix + c.textContent, li);
      li.querySelectorAll(":scope > img, :scope > p > img").forEach(img);
      li.querySelectorAll(":scope > ul, :scope > ol").forEach(walkList);
      n++;
    }
  };
  const walk = (el) => {
    for (const n of el.childNodes) {
      if (n.nodeType === 3) { line(n.textContent, null); continue; }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (/^(P|H[1-6])$/.test(tag)) {
        line(n.textContent, n);
        n.querySelectorAll("img").forEach(img);
      } else if (tag === "OL" || tag === "UL") {
        walkList(n);
      } else if (tag === "TABLE") {
        items.push({ kind: "extra", what: "Tabelle" });
        n.querySelectorAll("tr").forEach((tr) => {
          const cells = [...tr.children].map((td) => clean(td.textContent)).filter(Boolean);
          if (cells.length) items.push({ kind: "line", text: cells.join(" | "), id: n.getAttribute("id") || "" });
        });
      } else if (tag === "IMG") {
        img(n);
      } else if (tag === "OBJECT" || tag === "IFRAME") {
        items.push({ kind: "extra", what: "Anhang" });
      } else if (tag !== "BR" && tag !== "SCRIPT" && tag !== "STYLE") {
        walk(n);
      }
    }
  };
  walk(root);
  return items;
}

// Erkennt Aufgabenkennzeichen am Zeilenanfang: "1)", "(2)", "3a)", "B)", "c)"
const TASK_RE = /^\(?(\d{1,2}[a-zA-Z]?|[a-zA-Z])\)\s*(.*)$/;

// Prüft, ob eine Zeile ein Feedback-Feld einleitet: "Feedback:", "Feedback 2):", "Feedback b:" …
function matchMarker(line, base) {
  if (!norm(line).startsWith(base)) return null;
  const rest = clean(line).slice(base.length);
  const colon = rest.indexOf(":");
  if (colon >= 0 && colon <= 12) {
    return { label: rest.slice(0, colon).trim(), answer: rest.slice(colon + 1).trim() };
  }
  if (/^[\p{L}\p{N}]/u.test(rest)) return null;
  return { label: "", answer: rest.trim() };
}

const labelKey = (s) => (s || "").replace(/[()\s.]/g, "").toLowerCase();

// Position eines Textfelds auf der Seite (für die Reihenfolge von oben nach unten)
function posOf(el) {
  const st = el.getAttribute("style") || "";
  const get = (k) => {
    const m = new RegExp(`(?:^|;)\\s*${k}\\s*:\\s*(-?[\\d.]+)`).exec(st);
    return m ? parseFloat(m[1]) : 0;
  };
  return { top: get("top"), left: get("left") };
}

// Zerlegt eine Seite in Aufgaben. Eine Aufgabe beginnt mit ihrem Kennzeichen und reicht
// bis zum nächsten Kennzeichen oder bis zum Ende des Textfelds.
function parsePage(html, sel) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  let boxes = [...doc.body.children].filter((e) => e.tagName === "DIV");
  if (!boxes.length) boxes = [doc.body];
  boxes = boxes
    .map((el, i) => ({ el, i, ...posOf(el) }))
    .sort((a, b) => a.top - b.top || a.left - b.left || a.i - b.i)
    .map((b) => b.el);

  // Großbuchstaben (Teil A, B …) und Zahlen zählen immer als Aufgabe. Kleinbuchstaben nur,
  // wenn alle oder Unteraufgaben gewählt sind – sonst bleiben z. B. Auswahlmöglichkeiten
  // "a) … b) …" Teil der Aufgabe.
  const useLower = sel.all || sel.lower;
  const fields = [];
  const taskLines = [];
  const fbBoxes = [];
  let lastNum = "", lastUpper = "";

  for (const box of boxes) {
    const items = collect(box);
    const firstLine = items.find((x) => x.kind === "line");
    const boxId = box.getAttribute("id") || "";

    // Eigenes Feedback-Feld der Vorlage
    const fb = firstLine && matchMarker(firstLine.text, "feedback");
    if (fb) {
      const lines = items.filter((x) => x.kind === "line").map((x) => x.text);
      lines[0] = fb.answer;
      fbBoxes.push({ label: fb.label, id: boxId, lines: lines.filter(Boolean) });
      continue;
    }

    let cur = null;
    for (const it of items) {
      if (it.kind === "line") {
        const m = TASK_RE.exec(it.text);
        const isNum = m && /^\d/.test(m[1]);
        const isUpper = m && !isNum && m[1] !== m[1].toLowerCase();
        if (m && (isNum || isUpper || useLower)) {
          const own = m[1];
          let display = own;
          if (isNum) { lastNum = own; }
          else if (own === own.toLowerCase()) { display = (lastNum || "") + own; }
          else { lastUpper = own; }
          cur = {
            own, display, upper: lastUpper, isNum, boxId, lastId: it.id,
            raw: m[2] ? [m[2]] : [], imageSrcs: [], extras: [], existingFeedback: "", inFeedback: false,
          };
          fields.push(cur);
          continue;
        }
        if (!cur) { taskLines.push(it.text); continue; }
        if (it.id) cur.lastId = it.id;
        if (cur.inFeedback || isFeedbackLine(it.text)) {
          cur.inFeedback = true;
          cur.existingFeedback += (cur.existingFeedback ? "\n" : "") + it.text.replace(/^💬\s*feedback:\s*/i, "");
        } else {
          cur.raw.push(it.text);
        }
      } else if (cur) {
        if (it.id) cur.lastId = it.id;
        if (it.kind === "img") cur.imageSrcs.push(it.src);
        else if (!cur.extras.includes(it.what)) cur.extras.push(it.what);
      }
    }
  }

  // Doppelte Nummern (z. B. "A) 1) 2)  B) 1) 2)") mit dem Großbuchstaben davor unterscheiden
  const count = new Map();
  fields.forEach((f) => count.set(f.display, (count.get(f.display) || 0) + 1));
  fields.forEach((f) => { if (f.isNum && count.get(f.display) > 1 && f.upper) f.display = f.upper + f.display; });
  const seen = new Map();
  fields.forEach((f) => {
    const n = (seen.get(f.display) || 0) + 1;
    seen.set(f.display, n);
    f.key = n > 1 ? `${f.display}~${n}` : f.display;
    f.label = f.display;
    f.content = f.raw.slice();
    f.task = "";
    f.fbTarget = f.lastId || f.boxId;      // Standard: direkt unter die Antwort
    f.fbAction = f.lastId ? "insert" : "append";
    f.fbBox = null;
  });
  return { fields, fbBoxes, taskText: taskLines.join("\n").slice(0, 3000) };
}

// Aufgabentext der Vorlage erkennen: Zeilen am Anfang einer Aufgabe, die bei (fast) allen
// Schüler:innen gleich sind, sind der Aufgabentext – nicht die Antwort.
function stripTemplate(results) {
  const groups = new Map();
  results.forEach((r) => r.allFields.forEach((f) => {
    if (!groups.has(f.key)) groups.set(f.key, []);
    groups.get(f.key).push(f);
  }));
  for (const list of groups.values()) {
    list.forEach((f) => { f.content = f.raw.slice(); f.task = ""; });
    if (list.length < 3) continue;
    const need = Math.max(3, Math.ceil(list.length * 0.6));
    let active = list;
    const taskLines = [];
    for (let j = 0; active.length; j++) {
      const freq = new Map();
      active.forEach((f) => { const t = f.raw[j]; if (t !== undefined) { const k = norm(t); freq.set(k, (freq.get(k) || 0) + 1); } });
      let best = "", bestN = 0;
      freq.forEach((n, k) => { if (n > bestN) { best = k; bestN = n; } });
      if (bestN < need) break;
      // Wer direkt hinter den Aufgabentext geschrieben hat: nur den Teil danach als Antwort
      active.forEach((f) => {
        const t = f.raw[j];
        if (t !== undefined && norm(t) !== best && norm(t).startsWith(best + " ")) {
          f.content = [t.slice(best.length).trim()].concat(f.raw.slice(j + 1));
        }
      });
      active = active.filter((f) => f.raw[j] !== undefined && norm(f.raw[j]) === best);
      taskLines.push(active[0].raw[j]);
      active.forEach((f) => { f.content = f.raw.slice(j + 1); });
    }
    // Vorgegebene Zeilen wie "b) I ___ (go) …" mitten in der Aufgabe: kommen sie bei den
    // meisten gleich vor und beginnen mit einem Kennzeichen, gehören sie zur Vorlage.
    const freqAll = new Map();
    list.forEach((f) => new Set(f.raw.map(norm)).forEach((k) => freqAll.set(k, (freqAll.get(k) || 0) + 1)));
    const tmpl = new Set([...freqAll].filter(([k, n]) => n >= need && TASK_RE.test(k)).map(([k]) => k));
    if (tmpl.size) {
      list.forEach((f) => { f.content = f.content.filter((l) => !tmpl.has(norm(l))); });
      list[0].raw.forEach((l) => { if (tmpl.has(norm(l)) && !taskLines.includes(l)) taskLines.push(l); });
    }
    const task = taskLines.join("\n");
    list.forEach((f) => { f.task = task; });
  }
}

// Ordnet jeder Aufgabe ein Feedback-Feld der Vorlage zu:
//  "Feedback 2):" ↔ "2)"; sonst der Reihe nach; gibt es nur eins, teilen sich alle Aufgaben dieses Feld.
function assignFeedbackBoxes(fields, boxes) {
  boxes = (boxes || []).filter((b) => b.id);
  fields.forEach((f) => {
    f.fbTarget = f.lastId || f.boxId;
    f.fbAction = f.lastId ? "insert" : "append";
    f.fbBox = null;
  });
  if (!boxes.length || !fields.length) return;
  const used = new Set();
  const open = [];
  fields.forEach((f) => {
    const b = boxes.find((x) => !used.has(x) && x.label &&
      (labelKey(x.label) === labelKey(f.display) || labelKey(x.label) === labelKey(f.own)));
    if (b) { used.add(b); useBox(f, b, "dedicated"); } else open.push(f);
  });
  const rest = boxes.filter((b) => !used.has(b) && !b.label);
  if (!open.length) return;
  if (rest.length === open.length) {
    open.forEach((f, i) => useBox(f, rest[i], "dedicated"));
  } else if (rest.length === 1) {
    open.forEach((f) => useBox(f, rest[0], fields.length > 1 ? "shared" : "dedicated"));
  }
}

function useBox(f, box, kind) {
  f.fbTarget = box.id;
  f.fbAction = "append";
  f.fbBox = kind;
  const strip = (l) => l.replace(/^💬\s*feedback:\s*/i, "");
  if (kind === "dedicated") {
    f.existingFeedback = box.lines.map(strip).join("\n");
  } else {
    const prefix = labelKey(f.display);
    const own = box.lines.map((l) => TASK_RE.exec(clean(l))).filter((m) => m && labelKey(m[1]) === prefix);
    f.existingFeedback = own.map((m) => m[2]).join("\n");
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
  sel.innerHTML = `<option value="">Alle Aufgaben</option>` +
    [...keys].map(([k, d]) => `<option value="${esc(k)}">Aufgabe ${esc(d)})</option>`).join("");
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
    case "kein-feld": return "nicht gefunden";
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
        ${multi ? `<div class="flabel">Aufgabe ${esc(f.display)})</div>` : ""}
        ${f.text ? `<p class="answer">${esc(f.text)}</p>` : ""}
        ${(f.images || []).length ? `<div class="thumbs">${f.images.map((u) => `<a href="${esc(u)}" target="_blank" rel="noopener"><img src="${esc(u)}" alt="Bild von ${esc(r.name)}"></a>`).join("")}</div>` : ""}
        ${f.extras.length ? `<div class="note">📎 Enthält: ${esc(f.extras.join(", "))}</div>` : ""}
        ${f.status === "leer" && multi ? `<div class="empty">noch leer</div>` : ""}
        ${feedbackHtml(r, f)}
      </div>`).join("");

    const notes = [];
    if (v.status === "leer") notes.push("Falls mit dem Stift direkt auf die Seite geschrieben: Diese Handschrift wird hier nicht angezeigt – bitte Seite öffnen.");
    if (v.status === "kein-feld") notes.push("Keine der gewählten Aufgaben auf dieser Seite gefunden – evtl. Kennzeichen gelöscht oder verändert.");
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
  $("markerInput").value = LS.get("aue.tasks") || "";
  $("markerInput").addEventListener("change", () => { LS.set("aue.tasks", $("markerInput").value.trim()); reevaluate(); });
  $("markerInput").addEventListener("keydown", (e) => { if (e.key === "Enter") e.target.blur(); });
  $("taskChips").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-task]");
    if (b) toggleTask(b.dataset.task);
  });

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
