/*
 * Antwort-Übersicht – OneNote-Add-in für Kursnotizbücher
 *
 * Ablauf:
 *  1. Titel der gerade geöffneten Seite lesen.
 *  2. Alle Schüler-Abschnittsgruppen im Kursnotizbuch durchgehen
 *     (Gruppen, deren Name mit "_" beginnt – z. B. _Lehrkräfte, _Inhaltsbibliothek,
 *     _Platz für Zusammenarbeit – werden übersprungen).
 *  3. Bei jedem Schüler die Seite mit demselben Titel suchen.
 *  4. Auf dieser Seite das Textfeld (Container) finden, dessen erste Zeile
 *     mit der Markierung beginnt (Standard: "Answer:").
 *  5. Alles, was darin steht, als Antwort anzeigen.
 *
 * Es werden keine Daten an fremde Server geschickt – alles läuft im Browser.
 */
"use strict";

const SKIP_PREFIX = "_";
const STORAGE_KEY = "antwortUebersicht.marker";

let results = [];       // [{ name, status, text, ink, extras }]
let currentTitle = "";

const $ = (id) => document.getElementById(id);

Office.onReady((info) => {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) $("marker").value = saved;
  } catch (e) { /* Speicher nicht verfügbar – egal */ }

  $("load").addEventListener("click", loadAnswers);
  $("onlyMissing").addEventListener("change", render);

  if (info.host !== Office.HostType.OneNote) {
    setStatus("Dieses Add-in funktioniert nur in OneNote (im Web).", true);
    $("load").disabled = true;
  }
});

function norm(s) {
  return (s || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function setStatus(msg, isError) {
  const el = $("status");
  el.textContent = msg || "";
  el.classList.toggle("error", !!isError);
}

async function loadAnswers() {
  const markerRaw = $("marker").value.trim();
  if (!markerRaw) { setStatus("Bitte gib die erste Zeile des Antwortfelds an.", true); return; }
  try { localStorage.setItem(STORAGE_KEY, markerRaw); } catch (e) { /* egal */ }
  const marker = norm(markerRaw);

  $("load").disabled = true;
  setStatus("Lese Notizbuch …");
  $("summary").innerHTML = "";
  $("results").innerHTML = "";

  try {
    await OneNote.run(async (ctx) => {
      const app = ctx.application;
      const activePage = app.getActivePage();
      activePage.load("title");
      const groups = app.getActiveNotebook().sectionGroups;
      groups.load("items/name");
      await ctx.sync();

      currentTitle = activePage.title || "";
      const wanted = norm(currentTitle);
      if (!wanted) throw new Error("Die geöffnete Seite hat keinen Titel. Die Seiten werden über den Titel zugeordnet.");

      // 1) Schülergruppen
      const students = groups.items
        .filter((g) => !g.name.startsWith(SKIP_PREFIX))
        .map((g) => ({ name: g.name, group: g, sections: [], page: null, outlines: [] }));
      if (!students.length) {
        throw new Error("Keine Schüler-Abschnittsgruppen gefunden. Ist ein Kursnotizbuch geöffnet – und bist du als Lehrkraft angemeldet?");
      }

      // 2) Abschnitte (inkl. eine Ebene Unter-Abschnittsgruppen)
      students.forEach((s) => {
        s.group.sections.load("items/name");
        s.group.sectionGroups.load("items/name");
      });
      await ctx.sync();
      students.forEach((s) => {
        s.sections.push(...s.group.sections.items);
        s.group.sectionGroups.items.forEach((sg) => sg.sections.load("items/name"));
      });
      await ctx.sync();
      students.forEach((s) =>
        s.group.sectionGroups.items.forEach((sg) => s.sections.push(...sg.sections.items))
      );

      // 3) Seiten mit gleichem Titel suchen
      setStatus(`Suche Seite „${currentTitle}“ bei ${students.length} Schüler:innen …`);
      students.forEach((s) => s.sections.forEach((sec) => sec.pages.load("items/title")));
      await ctx.sync();
      students.forEach((s) => {
        for (const sec of s.sections) {
          const p = sec.pages.items.find((pg) => norm(pg.title) === wanted);
          if (p) { s.page = p; break; }
        }
        if (s.page) s.page.contents.load("items/type");
      });
      await ctx.sync();

      // 4) Textfelder (Outlines) und ihre Absätze laden
      setStatus("Lese Antwortfelder …");
      students.forEach((s) => {
        if (!s.page) return;
        s.outlines = s.page.contents.items
          .filter((c) => c.type === "Outline")
          .map((c) => c.outline);
        s.outlines.forEach((o) => o.paragraphs.load("items/type"));
      });
      await ctx.sync();

      // eine Ebene eingerückter Absätze (Aufzählungen etc.) mitnehmen
      const allParas = [];
      students.forEach((s) => s.outlines.forEach((o) => o.paragraphs.items.forEach((p) => {
        p.paragraphs.load("items/type");
      })));
      await ctx.sync();
      students.forEach((s) => s.outlines.forEach((o) => {
        o.flat = [];
        o.paragraphs.items.forEach((p) => {
          o.flat.push(p);
          p.paragraphs.items.forEach((child) => o.flat.push(child));
        });
        o.flat.forEach((p) => allParas.push(p));
      }));
      allParas.forEach((p) => { if (p.type === "RichText") p.richText.load("text"); });
      await ctx.sync();

      // 5) Auswerten
      const evaluated = students.map((s) => evaluate(s, marker));

      // 6) Bilder aus dem Antwortfeld holen (max. 4 pro Person)
      const pending = [];
      evaluated.forEach((r) => r.imgParas.slice(0, 4).forEach((p) => {
        pending.push({ r, data: p.image.getBase64Image() });
      }));
      if (pending.length) { setStatus("Lade Bilder …"); await ctx.sync(); }
      pending.forEach(({ r, data }) => { if (data.value) r.images.push("data:image/png;base64," + data.value); });

      results = evaluated.map(({ imgParas, ...rest }) => rest);
    });

    setStatus(`Stand: ${new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" })} Uhr`);
    render();
  } catch (err) {
    console.error(err);
    const detail = err && err.debugInfo ? ` (${err.debugInfo.message || err.debugInfo.code})` : "";
    setStatus("Fehler: " + (err.message || err) + detail, true);
  } finally {
    $("load").disabled = false;
  }
}

function evaluate(s, marker) {
  const r = { name: s.name, status: "", text: "", ink: false, extras: [], images: [], imgParas: [] };
  if (!s.page) { r.status = "keine-seite"; return r; }

  // Textfeld suchen, dessen erste Textzeile mit der Markierung beginnt
  let box = null, markerPara = null;
  for (const o of s.outlines) {
    const firstText = o.flat.find((p) => p.type === "RichText");
    if (firstText && norm(firstText.richText.text).startsWith(marker)) {
      box = o; markerPara = firstText; break;
    }
  }
  if (!box) { r.status = "kein-feld"; return r; }

  const lines = [];
  for (const p of box.flat) {
    if (p === markerPara) {
      // Text, der in derselben Zeile hinter "Answer:" steht
      const rest = restAfterMarker(p.richText.text, marker);
      if (rest) lines.push(rest);
      continue;
    }
    switch (p.type) {
      case "RichText": {
        const t = (p.richText.text || "").trim();
        if (t) lines.push(t);
        break;
      }
      case "Ink":   r.ink = true; break;
      case "Image": r.imgParas.push(p); break;
      case "Table": r.extras.push("Tabelle"); break;
      default: break;
    }
  }
  r.text = lines.join("\n");
  r.status = (r.text || r.ink || r.extras.length || r.imgParas.length) ? "ok" : "leer";
  return r;
}

function restAfterMarker(text, marker) {
  const t = (text || "").replace(/\s+/g, " ").trim();
  return t.length > marker.length ? t.slice(marker.length).trim() : "";
}

const LABELS = {
  "ok": "beantwortet",
  "leer": "noch nichts",
  "kein-feld": "Feld fehlt",
  "keine-seite": "Seite fehlt",
};

function render() {
  const box = $("results");
  const sum = $("summary");
  if (!results.length) { box.innerHTML = ""; sum.innerHTML = ""; return; }

  const answered = results.filter((r) => r.status === "ok").length;
  const pct = Math.round((answered / results.length) * 100);
  sum.innerHTML =
    `${escapeHtml(currentTitle)}<br>${answered} von ${results.length} haben geantwortet` +
    `<div class="bar"><span style="width:${pct}%"></span></div>`;

  const onlyMissing = $("onlyMissing").checked;
  const shown = onlyMissing ? results.filter((r) => r.status !== "ok") : results;

  if (!shown.length) {
    box.innerHTML = `<p class="hint">Alle haben geantwortet. 🎉</p>`;
    return;
  }

  box.innerHTML = shown.map((r) => {
    const notes = [];
    if (r.ink) notes.push("✍️ Enthält Handschrift/Zeichnung (hier nicht lesbar)");
    if (r.extras.length) notes.push("📎 Enthält: " + [...new Set(r.extras)].join(", "));
    if (r.status === "kein-feld") notes.push("Das Antwortfeld wurde nicht gefunden – evtl. gelöscht oder die erste Zeile verändert.");
    if (r.status === "keine-seite") notes.push("Diese Seite wurde bei dieser Person nicht gefunden.");

    return `<article class="card ${r.status}">
      <h2><span>${escapeHtml(r.name)}</span><span class="badge">${LABELS[r.status]}</span></h2>
      ${r.text ? `<p class="answer">${escapeHtml(r.text)}</p>` : ""}
      ${r.images.length ? `<div class="thumbs">${r.images.map((u) => `<img src="${escapeHtml(u)}" alt="Bild">`).join("")}</div>` : ""}
      ${notes.map((n) => `<div class="note">${escapeHtml(n)}</div>`).join("")}
    </article>`;
  }).join("");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}
