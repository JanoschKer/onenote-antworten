/*
 * KI-Feedback für die Antwort-Übersicht
 *
 * - Erstellt Feedback-Vorschläge mit Azure OpenAI oder Claude (Anthropic).
 * - Gesendet werden nur: Aufgabentext der Seite, deine Hinweise, die Antwort.
 *   Namen werden NICHT gesendet.
 * - Der Schlüssel wird nur auf diesem Gerät gespeichert (localStorage).
 * - "In OneNote eintragen" hängt das Feedback unten an das Antwortfeld an.
 */
"use strict";

const FB_PREFIX = "💬 Feedback:";
const fbState = new Map();   // Schlüssel -> { state: loading|ready|error|saved, text, error }
const fbIndex = new Map();   // Schlüssel -> { r, f } (wird bei jedem Zeichnen neu gefüllt)

function isFeedbackLine(line) {
  return /^💬\s*feedback:/i.test(clean(line));
}

/* ---------------- Einstellungen ---------------- */

const AI_DEFAULTS = {
  provider: "anthropic",
  azEndpoint: "", azDeployment: "", azKey: "",
  anKey: "", anModel: "claude-sonnet-5-5",
  lang: "auto",
};

function getAI() {
  try { return { ...AI_DEFAULTS, ...JSON.parse(LS.get("aue.ai") || "{}") }; }
  catch (e) { return { ...AI_DEFAULTS }; }
}

function aiReady(s) {
  return s.provider === "azure"
    ? !!(s.azEndpoint && s.azDeployment && s.azKey)
    : !!(s.anKey && s.anModel);
}

function critKey() { return "aue.crit." + norm($("titleInput").value); }

function initAISettings() {
  const s = getAI();
  ["provider", "azEndpoint", "azDeployment", "azKey", "anKey", "anModel", "lang"].forEach((k) => {
    const el = $("ai_" + k);
    el.value = s[k];
    el.addEventListener("input", saveAI);
    el.addEventListener("change", saveAI);
  });
  $("ai_crit").value = LS.get(critKey()) || "";
  $("ai_crit").addEventListener("input", () => LS.set(critKey(), $("ai_crit").value));
  $("titleInput").addEventListener("change", () => { $("ai_crit").value = LS.get(critKey()) || ""; });
  showProviderFields();
  updateAIHint();

  $("results").addEventListener("click", onFeedbackClick);
  $("results").addEventListener("input", (e) => {
    const k = e.target.dataset && e.target.dataset.k;
    if (e.target.tagName === "TEXTAREA" && k && fbState.has(k)) fbState.get(k).text = e.target.value;
  });
  $("fbAllBtn").addEventListener("click", feedbackForAll);
}

function saveAI() {
  const s = {};
  ["provider", "azEndpoint", "azDeployment", "azKey", "anKey", "anModel", "lang"].forEach((k) => {
    s[k] = $("ai_" + k).value.trim();
  });
  LS.set("aue.ai", JSON.stringify(s));
  showProviderFields();
  updateAIHint();
}

function showProviderFields() {
  const az = $("ai_provider").value === "azure";
  $("ai_azure").hidden = !az;
  $("ai_anthropic").hidden = az;
}

function updateAIHint() {
  $("ai_state").textContent = aiReady(getAI()) ? "eingerichtet ✓" : "noch nicht eingerichtet";
}

/* ---------------- Anzeige im Kärtchen ---------------- */

function fbKeyOf(r, f) { return r.page.id + "|" + f.key; }

function feedbackHtml(r, f) {
  if (!r.page) return "";
  const k = fbKeyOf(r, f);
  fbIndex.set(k, { r, f });
  const st = fbState.get(k);
  const ek = esc(k);
  let h = "";

  if (f.existingFeedback && !(st && st.state === "saved")) {
    h += `<div class="fb-existing">💬 Eingetragenes Feedback: ${esc(f.existingFeedback)}</div>`;
  }
  if (f.status !== "ok" || !f.text) return h;

  if (!st) {
    return h + `<button type="button" class="fb-btn" data-act="gen" data-k="${ek}">✨ ${f.existingFeedback ? "Neues " : ""}KI-Feedback</button>`;
  }
  if (st.state === "loading") return h + `<div class="note">✨ Feedback wird erstellt …</div>`;
  if (st.state === "error") {
    return h + `<div class="fb-error">${esc(st.error)}</div>
      <button type="button" class="fb-btn" data-act="gen" data-k="${ek}">Erneut versuchen</button>`;
  }
  return h + `<div class="fb-box">
      <textarea data-k="${ek}" rows="4" ${st.state === "saved" ? "readonly" : ""}>${esc(st.text)}</textarea>
      <div class="fb-actions">
        ${st.state === "saved"
          ? `<span class="fb-saved">✓ In OneNote eingetragen</span>`
          : `<button type="button" class="fb-btn primary-sm" data-act="save" data-k="${ek}">In OneNote eintragen</button>
             <button type="button" class="fb-btn" data-act="gen" data-k="${ek}">Neu erstellen</button>
             <button type="button" class="fb-btn" data-act="discard" data-k="${ek}">Verwerfen</button>`}
      </div>
    </div>`;
}

function onFeedbackClick(e) {
  const btn = e.target.closest("button[data-act]");
  if (!btn) return;
  const k = btn.dataset.k;
  if (btn.dataset.act === "gen") generateFeedback(k);
  else if (btn.dataset.act === "save") saveFeedback(k);
  else if (btn.dataset.act === "discard") { fbState.delete(k); render(); }
}

/* ---------------- Feedback erstellen ---------------- */

function langText(code) {
  return code === "de" ? "Deutsch" : code === "en" ? "Englisch" : "dieselbe Sprache, in der die Aufgabe gestellt ist";
}

function buildPrompt(r, f, s) {
  const system =
`Du unterstützt eine Lehrkraft dabei, Feedback zu Schülerantworten zu formulieren. Das Feedback richtet sich direkt an die Schülerin bzw. den Schüler (du-Form; auf Englisch "you").
Regeln:
- 2 bis 4 kurze Sätze, freundlich, ehrlich und konkret.
- Nenne zuerst, was gelungen ist, dann genau einen konkreten nächsten Schritt zur Verbesserung.
- Verrate nicht die vollständige Musterlösung.
- Keine Noten oder Punkte, außer die Hinweise der Lehrkraft verlangen es ausdrücklich.
- Ist die Antwort leer, unverständlich oder geht an der Aufgabe vorbei, sag das freundlich und gib einen Tipp.
- Gib nur den Feedbacktext aus – ohne Anrede mit Namen, ohne Überschrift, ohne Anführungszeichen.
- Sprache des Feedbacks: ${langText(s.lang)}.`;

  const crit = (LS.get(critKey()) || "").trim();
  const user =
`Seitentitel: ${state.title}

Aufgabentext auf der Seite:
${r.taskText || "(kein weiterer Text auf der Seite)"}

Hinweise / Erwartungshorizont der Lehrkraft:
${crit || "(keine)"}

${r.fields.length > 1 ? `Es geht um Antwortfeld ${f.display}.\n\n` : ""}Antwort der Schülerin / des Schülers:
${f.text}`;
  return { system, user };
}

// Ruft die KI auf – mit Wiederholung bei Überlastung und Prüfung auf abgeschnittene Antworten
async function callAI(system, user) {
  let budget = 1500;
  let busy = 0, redo = 0;
  for (;;) {
    let res;
    try {
      res = await callAIOnce(system, user, budget);
    } catch (e) {
      // 429 = zu viele Anfragen, 529/503 = Dienst überlastet, 0 = Netzwerk → kurz warten, erneut
      if ([0, 429, 500, 502, 503, 504, 529].includes(e.status) && busy < 4) {
        const wait = e.retryAfter || Math.min(30, 3 * 2 ** busy);
        busy++;
        setStatus(`KI ist gerade ausgelastet – neuer Versuch in ${wait} s …`);
        await sleep(wait * 1000);
        continue;
      }
      throw e;
    }
    if (res.truncated && redo < 2) { redo++; budget *= 2; continue; }   // abgeschnitten → mit mehr Platz neu
    if (!res.text) {
      if (res.reason === "content_filter") throw new Error("Der Inhaltsfilter von Azure hat die Antwort blockiert.");
      if (redo < 2) { redo++; continue; }
      throw new Error(`Die KI hat keinen Text geliefert (Grund: ${res.reason || "unbekannt"}).`);
    }
    if (res.truncated) res.text += " …";
    return res;
  }
}

function aiError(prefix, r, j) {
  const err = new Error(prefix + ((j && j.error && j.error.message) || `HTTP ${r.status}`));
  err.status = r.status;
  const ra = Number(r.headers.get("retry-after"));
  if (ra > 0) err.retryAfter = Math.min(ra, 60);
  return err;
}

async function safeFetch(url, opt) {
  try { return await fetch(url, opt); }
  catch (e) { const err = new Error("Keine Verbindung zur KI (" + e.message + ")"); err.status = 0; throw err; }
}

async function callAIOnce(system, user, budget) {
  const s = getAI();
  if (s.provider === "azure") {
    const url = `${s.azEndpoint.replace(/\/+$/, "")}/openai/deployments/${encodeURIComponent(s.azDeployment)}` +
      `/chat/completions?api-version=2024-10-21`;
    const r = await safeFetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "api-key": s.azKey },
      body: JSON.stringify({ messages: [{ role: "system", content: system }, { role: "user", content: user }] }),
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw aiError("Azure OpenAI: ", r, j);
    const c = (j.choices && j.choices[0]) || {};
    let content = c.message && c.message.content;
    if (Array.isArray(content)) content = content.map((x) => x.text || "").join("");
    return { text: (content || "").trim(), truncated: c.finish_reason === "length", reason: c.finish_reason };
  }
  const r = await safeFetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": s.anKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: s.anModel,
      max_tokens: budget,
      system,
      messages: [{ role: "user", content: user }],
    }),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw aiError("Claude: ", r, j);
  // Nur Textblöcke verwenden (eventuelle "Denk"-Blöcke ignorieren)
  const text = (j.content || []).filter((b) => b.type === "text").map((b) => b.text).join("").trim();
  return { text, truncated: j.stop_reason === "max_tokens", reason: j.stop_reason };
}

async function generateFeedback(k) {
  const item = fbIndex.get(k);
  if (!item) return;
  const s = getAI();
  if (!aiReady(s)) {
    $("aiPanel").open = true;
    $("aiPanel").scrollIntoView({ behavior: "smooth" });
    setStatus("Bitte zuerst unter „KI-Feedback einrichten“ einen Anbieter und Schlüssel eintragen.", true);
    return;
  }
  fbState.set(k, { state: "loading" });
  render();
  try {
    const { system, user } = buildPrompt(item.r, item.f, s);
    const { text } = await callAI(system, user);
    fbState.set(k, { state: "ready", text });
  } catch (e) {
    console.error(e);
    fbState.set(k, { state: "error", error: e.message || String(e) });
  }
  render();
}

async function feedbackForAll() {
  if (!aiReady(getAI())) {
    $("aiPanel").open = true;
    setStatus("Bitte zuerst unter „KI-Feedback einrichten“ einen Anbieter und Schlüssel eintragen.", true);
    return;
  }
  const keys = [...fbIndex.entries()]
    .filter(([k, { f }]) => f.status === "ok" && f.text && !f.existingFeedback && !fbState.has(k))
    .map(([k]) => k);
  if (!keys.length) { setStatus("Keine offenen Antworten ohne Feedback in der aktuellen Ansicht."); return; }
  $("fbAllBtn").disabled = true;
  setStatus(`Erstelle Feedback für ${keys.length} Antworten …`);
  await pool(keys, 2, (k) => generateFeedback(k));
  $("fbAllBtn").disabled = false;
  setStatus(`Feedback für ${keys.length} Antworten erstellt – bitte prüfen.`);
}

/* ---------------- In OneNote eintragen ---------------- */

async function saveFeedback(k) {
  const item = fbIndex.get(k);
  const st = fbState.get(k);
  if (!item || !st) return;
  if (!CFG.feedbackSchreiben) {
    setStatus("Eintragen ist noch nicht freigeschaltet – siehe README, Teil C (Schreibrecht + config.js).", true);
    return;
  }
  const { r, f } = item;
  if (!f.targetId) { setStatus("Das Antwortfeld konnte nicht eindeutig bestimmt werden. Bitte neu laden.", true); return; }

  const lines = st.text.split(/\n+/).map((l) => l.trim()).filter(Boolean);
  const color = "color:#7719aa";
  const html = lines.map((l, i) => i === 0
    ? `<p style="${color}">💬 <b>Feedback:</b> ${esc(l)}</p>`
    : `<p style="${color}">${esc(l)}</p>`).join("");

  try {
    setStatus(`Trage Feedback bei ${r.name} ein …`);
    await graph(`${state.base}/pages/${r.page.id}/content`, "none", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify([{ target: f.targetId, action: "append", content: html }]),
    });
    st.state = "saved";
    f.existingFeedback = st.text;
    setStatus(`Feedback bei ${r.name} eingetragen.`);
  } catch (e) {
    let msg = e.message || String(e);
    if (e.status === 403) msg = "Keine Schreibberechtigung. Wurde Notes.ReadWrite.All freigegeben und hast du dich danach neu angemeldet?";
    setStatus("Eintragen fehlgeschlagen: " + msg, true);
  }
  render();
}
