# Elternbrief (eigenes Tool)

**Stand:** 2026-11-18  
**Status:** Upgrade Welle A–C + SPO Live Sync S0–S3 (Familie `ablauf-kommunikation`)  
**Familie:** `ablauf-kommunikation` (Absender-Inbox + Public Magic-Link)  
**Roadmap:** [`schul-apps-roadmap-prozesse.md`](./schul-apps-roadmap-prozesse.md) · Vorlage: [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md) · SoR: [`spo-live-sync.md`](./spo-live-sync.md)  
**Lizenzpaket:** `ablauf` (Eltern & Ablauf)

---

## Ziel

Digitale **Elternbriefe** als eigenes Werkzeug (nicht nur Einsammler-Modus): Vorlagen, Text-Editor, optionaler Anhang, Zielklassen aus Stammdaten, persönlicher Leselink, Lesebestätigung und Übersicht „wer hat gelesen“.

Versandkanal ist bewusst der **Link** (mailto öffnet nur das lokale Mailprogramm). Automatischer Serienversand (Graph/PA) folgt später.

## Rollen (kurz)

| Schicht | Wert |
|---------|------|
| Portal-Persona | Lehrkraft (Haupttool); IT für `sharepoint-liste-elternbrief` |
| App-Rollen | Lehrkraft = nur eigene (`erstellerEmail`); Admin (Entra + Legacy-Mails) optional „Alle“ |
| Empfänger | Eltern ohne Microsoft-Login über Magic-Link |

## Abgrenzung zum Einsammler

| | Elternbrief | Einsammler |
|--|-------------|------------|
| Fokus | Informieren + gelesen? | Einwilligung / Abgabe |
| Inhalt | Vorlagen + längerer Text + Anhang | Kurzer Vorgangstext |
| Antwort | Lesebestätigung | Zustimmen / Ablehnen / Datei |

Einwilligungen (Foto, Exkursion) bleiben im **Einsammler**. Elternbrief kann darauf verweisen.

## Funktionen

- Vorlagen: Schulbeginn, Schularbeitenplan, Ausflug-Info, Elternabend, Gesundheit, frei  
- Editor: Betreff, Inhalt, Vorschau  
- Klassen-Chips aus Stammdaten (+ optional Live-Hydrate über Tool-Site, local-first bei 403)  
- Absender: eigenes Konto **oder** zentrales Postfach (Einstellungen / Stammdaten-Mailbox)  
- Optionaler Anhang (IndexedDB, max. 6 MB) – Download auf der öffentlichen Seite  
- Status: Entwurf → gesendet → geschlossen  
- Persönliche Links (`elternbrief-lesen.html#p=…`, HMAC-signiert; Legacy-`?p=` wird noch gelesen)  
- Lesebestätigung + Bestätigungscode / lokaler Inbox-Merge  
- Lehrkraft sieht nur eigene Briefe; Admin: „Alle“  
- SharePoint-Listen-Setup (Site über `resolveSchoolToolSiteUrl('elternbrief')`)  
- **SPO Live Sync (SoR):** `elternbrief-graph.js` + `elternbrief-sync.js` – Pull beim Start, Push bei Speichern/Status/Löschen; Demo-Seeds nicht pushen; 403 wipe’t Cache nicht  
- Public Magic-Link bestätigt lokal (Inbox/Code); Lehrkraft-Merge pusht Lesestatus nach SPO  
- Calm Absender-UI: Action-Banner (Entwurf→Freigeben / Links teilen), Quiet Empty, Demo nur Admin oder `?demoRole=1`

## Workflow

1. Anmelden  
2. Vorlage wählen / Text schreiben / Anhang  
3. Klassen wählen → Empfänger übernehmen  
4. Speichern → **Senden freigeben**  
5. Links kopieren oder „Mail vorbereiten“  
6. Eltern lesen (+ Anhang) und bestätigen  
7. Fortschritt in Übersicht / Empfänger-Tabelle  

## Tool-IDs

- `elternbrief`
- `sharepoint-liste-elternbrief`

## Site-Key

- Settings: `ms365-elternbrief-sp-settings-v1.siteUrl`
- Shared Kind: `elternbrief` in `school-tool-sites`
- Mailbox: `ms365-elternbrief-settings-v1` / Stammdaten `schoolMailboxes.elternbriefZentral`

## Dateien

- UI: `tools/elternbrief.html`, `tools/elternbrief-lesen.html`  
- Module: `src/tools/elternbrief/*` (`elternbrief-graph.js`, `elternbrief-sync.js`)  
- SP: `tools/sharepoint-liste-elternbrief.html`  
- Stammdaten-Reuse: `src/tools/einsammler/einsammler-stammdaten.js` (`siteKind: 'elternbrief'`)

## Smoke (SPO)

1. IT: Listen anlegen (`sharepoint-liste-elternbrief`) – `Elternbriefe` + `ElternbriefLesebestaetigungen`  
2. Lehrkraft: Brief speichern → Item in SPO-Liste sichtbar (`BriefId` / Betreff)  
3. Senden freigeben → Status in SPO `gesendet`; Empfänger-Zeilen in Lesebestätigungs-Liste  
4. Admin: Einstellungen → „Von SharePoint laden“; Sync-Meta `live`  
5. Demo laden (`?demoRole=1` oder Admin) → bleibt lokal, kein Push  
6. 403/keine Rechte → ruhiger Hinweis, lokaler Stand bleibt  
