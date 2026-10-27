# Antwort-Übersicht für OneNote-Kursnotizbücher

Zeigt dir die Antworten aller Schülerinnen und Schüler zu einer verteilten Seite – und markiert, wer noch nichts geschrieben hat.

Es gibt zwei Varianten im selben Paket:

| | **A: Add-in** (Hauptordner) | **B: Web-Übersicht** (Ordner `web/`) |
|---|---|---|
| Läuft in | OneNote für das Web (Seitenbereich) | jedem Browser – auch auf iPad/Android neben der OneNote-App |
| Einrichtung | Dateien hochladen + Manifest | Dateien hochladen + einmalige App-Registrierung bei Microsoft (meist durch die Schul-IT) |

Für beide gilt Schritt 1 (Seite vorbereiten) und Schritt 2 (Dateien auf GitHub Pages veröffentlichen). Danach weiter mit Teil A (Schritt 3–4) oder Teil B.

---

## 1. Seite vorbereiten (einmal pro Aufgabe)

1. Erstelle die Aufgabenseite in deinem Kursnotizbuch.
2. Klicke an eine freie Stelle und schreibe als **erste Zeile** dieses Textfelds:
   `Answer:`
   Darunter lässt du Platz. Die Schüler schreiben **in genau dieses Textfeld** (unter „Answer:“).
3. Verteile die Seite wie gewohnt (Kursnotizbuch → Seite verteilen).

Tipp für die Schüler: „Klickt direkt unter ‚Answer:‘ und schreibt dort weiter.“ Wer irgendwo anders auf die Seite klickt, erzeugt ein neues Textfeld – das wird nicht mitgezählt.

Du kannst statt `Answer:` auch eine andere Markierung nehmen – trage sie dann im Add-in ein.

### Handschriftliche Antworten

Freie Stift-Tinte auf der Seite kann von außen nicht zuverlässig ausgelesen werden (Microsoft bietet dafür nur eine Beta-Schnittstelle). Diese Wege funktionieren:

- **Bild ins Antwortfeld:** Die Schüler:innen schreiben auf Papier (oder in einer Zeichen-App), machen ein Foto bzw. einen Screenshot und fügen es **direkt unter „Answer:“** ein (Einfügen → Bild / Kamera). Das Bild erscheint als Vorschau in der Übersicht; antippen vergrößert es.
- **Stift wird zu Text:** Auf dem iPad mit Apple Pencil direkt ins Textfeld schreiben („Kritzeln“/Scribble wandelt die Schrift beim Schreiben in Text um). In OneNote auch: Zeichnen → „Freihand in Text“. Dann erscheint die Antwort ganz normal als Text.

Schreibt jemand mit dem Stift frei auf die Seite, zeigt die Übersicht „noch nichts“ – über „In OneNote öffnen“ kommst du direkt hin.

## 2. Dateien veröffentlichen (einmalig, ca. 10 Minuten)

Office-Add-ins müssen über eine HTTPS-Adresse erreichbar sein. Am einfachsten kostenlos mit **GitHub Pages**:

1. Konto auf github.com anlegen.
2. Neues **öffentliches** Repository namens `onenote-antworten` erstellen.
3. Alle Dateien aus diesem Ordner hochladen („Add file → Upload files“), inkl. Ordner `assets`.
4. Repository → **Settings → Pages** → Source: „Deploy from a branch“, Branch `main`, Ordner `/ (root)` → Save.
5. Nach 1–2 Minuten ist die Seite unter `https://DEINNAME.github.io/onenote-antworten/` erreichbar.
6. Öffne `manifest.xml` (z. B. im Editor) und ersetze **überall** `DEIN-NAME` durch deinen GitHub-Benutzernamen. Speichern.

(Die Manifest-Datei selbst muss nicht hochgeladen werden, schadet aber nicht.)

Datenschutz: Die hochgeladenen Dateien enthalten nur den Programmcode. Die Schülerantworten werden ausschließlich in deinem Browser gelesen und nirgendwohin gesendet.

## 3. Add-in in OneNote einbinden

1. Kursnotizbuch in **OneNote für das Web** öffnen.
2. **Einfügen → Office-Add-ins** (bzw. „Add-ins“) → **Mein Add-in hochladen** → `manifest.xml` auswählen.
3. Auf der Registerkarte **Start** erscheint der Knopf **Antwort-Übersicht**.

Falls „Mein Add-in hochladen“ fehlt: Dann hat deine Schul-IT das Hochladen eigener Add-ins gesperrt. Gib ihr die `manifest.xml` – sie kann das Add-in zentral im Microsoft-365-Admin-Center bereitstellen.

## 4. Benutzen

1. Öffne die verteilte Seite – z. B. bei irgendeinem Schüler oder deine eigene Kopie. Wichtig ist nur der **Seitentitel**.
2. **Antwort-Übersicht** → **Antworten laden**.
3. Du siehst:
   - 🟢 **beantwortet** – mit dem Text der Antwort
   - 🔴 **noch nichts** – Antwortfeld leer
   - 🟡 **Feld fehlt** – Antwortfeld gelöscht oder erste Zeile verändert
   - 🟡 **Seite fehlt** – Seite bei dieser Person nicht vorhanden (z. B. umbenannt)
4. „Nur fehlende Antworten zeigen“ blendet alle Erledigten aus.
5. Ziehst du den Seitenbereich breiter, stehen die Antworten nebeneinander.

Zum Aktualisieren einfach erneut auf **Antworten laden** klicken.

## Grenzen

- Handschrift (Stift/Ink) kann nicht als Text gelesen werden – das Add-in zeigt dann „Enthält Handschrift“ an.
- Bilder und Tabellen im Antwortfeld werden nur als Hinweis angezeigt.
- Die Zuordnung läuft über den Seitentitel. Benennt ein Schüler die Seite um, gilt sie als „Seite fehlt“.
- Abschnittsgruppen, deren Name mit `_` beginnt (Lehrkräfte-Bereich, Inhaltsbibliothek, Zusammenarbeit), werden übersprungen.
- Bei sehr großen Klassen/Notizbüchern kann das Laden einige Sekunden dauern.

## Dateien

| Datei | Zweck |
|---|---|
| `manifest.xml` | Beschreibt das Add-in für OneNote (hier die Adresse eintragen) |
| `taskpane.html` | Oberfläche im Seitenbereich |
| `taskpane.js` | Liest das Notizbuch aus und baut die Übersicht |
| `taskpane.css` | Gestaltung |
| `assets/` | Symbole |

---

# Teil B: Web-Übersicht für Tablet & Browser (Ordner `web/`)

Eine eigenständige Webseite, die dein Kursnotizbuch über Microsoft 365 **nur liest**. Du öffnest sie im Tablet-Browser (Safari, Chrome, Edge) – z. B. in geteilter Ansicht neben der OneNote-App. Jede Karte hat einen Link **„In OneNote öffnen“**, der direkt zur Seite der Person in der App springt.

Funktioniert mit Kursnotizbüchern in deinem OneDrive **und** mit Kursnotizbüchern aus Microsoft Teams.

## B1. App bei Microsoft registrieren (einmalig, ca. 10 Minuten)

Damit die Seite auf dein Notizbuch zugreifen darf, braucht sie eine „App-Registrierung“ im Microsoft-Konto deiner Schule. In vielen Schulen darf das nur die IT – dann gib ihr einfach diesen Abschnitt.

1. Öffne **entra.microsoft.com** → **Anwendungen → App-Registrierungen → Neue Registrierung**.
2. Name: `Antwort-Übersicht`
   Kontotypen: **Nur Konten in diesem Organisationsverzeichnis**
   Umleitungs-URI: Plattform **Single-Page-Anwendung (SPA)**,
   Adresse: `https://DEINNAME.github.io/onenote-antworten/web/` (mit Schrägstrich am Ende!)
   → **Registrieren**.
3. Auf der Übersichtsseite notieren:
   - **Anwendungs-ID (Client)**
   - **Verzeichnis-ID (Mandant)**
4. **API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen** und auswählen:
   - `Notes.Read.All` (Notizbücher lesen)
   - `Team.ReadBasic.All` (Liste deiner Teams, für Teams-Kursnotizbücher)
   - `User.Read` ist schon vorhanden.
5. Falls eure Schule Zustimmungen durch Nutzer sperrt: **„Administratorzustimmung erteilen“** (kann nur die IT).

Die App bekommt nur **Leserechte** und nur im Namen der angemeldeten Person – sie sieht also genau das, was du in OneNote ohnehin sehen kannst.

## B2. Einrichten

1. Öffne `web/config.js` und trage die beiden IDs aus B1 ein.
2. Lade den Ordner `web` mit ins GitHub-Repository (siehe Schritt 2 oben).
3. Öffne auf dem Tablet `https://DEINNAME.github.io/onenote-antworten/web/`
   Tipp: Über „Teilen → Zum Home-Bildschirm“ bekommst du ein App-Symbol.

## B3. Benutzen

1. **Mit Microsoft anmelden** (einmal; die Anmeldung bleibt ca. einen Tag erhalten).
2. **Ort** wählen (Mein OneDrive oder dein Team) → **Kursnotizbuch** wählen.
3. **Seitentitel** eintippen oder aus den Vorschlägen wählen → **Antworten laden**.
4. Optional: **Jede Minute aktualisieren** – praktisch während der Arbeitsphase.

## Grenzen der Web-Übersicht

- **Handschrift (Stift) wird von Microsoft nicht als Text ausgeliefert.** Ein nur handschriftlich ausgefülltes Feld erscheint als „noch nichts“ – dann über „In OneNote öffnen“ nachsehen.
- Die Zuordnung läuft über den Seitentitel (wie beim Add-in).
- Bei großen Klassen kann das erste Laden 10–30 Sekunden dauern; Microsoft begrenzt die Anzahl der Abfragen pro Minute.
- Datenschutz: Es gibt keinen eigenen Server. Die Seite holt die Daten direkt bei Microsoft und zeigt sie nur in deinem Browser an.
