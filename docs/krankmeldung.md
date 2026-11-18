# Krankmeldung (Eltern & Ablauf)

**Stand:** 2026-11-18  
**Status:** Upgrade Welle A–C + SPO Live Sync S0–S3 (Familie `ablauf-kommunikation`)  
**Familie:** `ablauf-kommunikation` (Sekretariats-Inbox + Public-Formular)  
**SoR:** SharePoint-Liste „Krankmeldungen“ (`krankmeldung-graph` / `krankmeldung-sync`)  
**Roadmap:** [`schul-apps-roadmap-prozesse.md`](./schul-apps-roadmap-prozesse.md)  
**Lizenzpaket:** `ablauf`  
**Upgrade-Vorlage:** [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md)

---

## Ziel

Eltern melden Kinder **krank** (Von/Bis, Kategorie). Sekretariat sieht eine **Inbox** (Tag/Woche/Klasse) – ohne Untis-Klassenbuch zu ersetzen.

## Rollen (kurz)

| Ebene | Inhalt |
|-------|--------|
| Portal-Personas | Lehrer/Sekretariat (Inbox), IT (Liste), Eltern (öffentliches Formular ohne Login) |
| App-Rollen | Admin/Sekretariat (Entra) · KV (Klassen-Scope) · Eltern (Public) |
| Nicht | Genehmigungsfluss, Bookings, Planer-Sidebar erzwingen |

## Funktionen

- Öffentliches Formular `tools/krankmeldung-formular.html` mit **Klassenauswahl** (Chips)
- Optionaler / Pflicht-**Anhang** (PDF/Foto/Word, IndexedDB; SharePoint-URL vorbereitet)
- Klassen für Eltern: Einstellungen „aus Stammdaten übernehmen“ + Link mit `?klassen=…`
- Admin-Inbox `tools/krankmeldung.html` (Filter Tag/Woche/Klasse, Zur Kenntnis mit **Actor-E-Mail**, Archiv, Anhang öffnen)
- Stammdaten optional live hydratisieren (`siteKind: 'krankmeldung'`), local-first bei 403
- Calm Action-Banner (Klassen veröffentlichen / heutige Neu / Link teilen); Quiet Empty ohne gestrichelte Lehrer-Panels
- Manuelle Erfassung mit denselben Klassen-Chips
- Sonderfälle: max. Zeitraum 30 Tage, Von nicht zu weit in Vergangenheit/Zukunft, Duplikat-Hinweis
- Demo **opt-in** (Admin-Button oder `?demo=1` / `?demoRole=1`) – local-first, Seed-Tag, Reset nur Demo
- Optional Remote-API (gleiche Function App wie Einsammler); API-Fehler soft (lokale Inbox bleibt)
- **SPO Live Sync (SoR):** `krankmeldung-graph.js` + `krankmeldung-sync.js` – Admin Pull/Push, `spoItemId`, 403 wipe’t nicht, Demo-Seeds nicht pushen
- Public-Formular: lokal + optional Remote; SPO-Push wenn Graph-Hilfen/Session (Kiosk); sonst Sekretariat nach API-Übernahme
- SharePoint-Liste anlegen; Site über `school-tool-sites` (`krankmeldung`) / Prefill aus Planer-Default

## Klassen für Eltern (wichtig)

Eltern haben keine Stammdaten im Browser. Deshalb:

1. Unter **Einstellungen** Klassen aus Stammdaten übernehmen (wird gespeichert)
2. **Formular-Link kopieren** – enthält `?klassen=1A;1B;…`
3. Diesen Link an Eltern weitergeben (Website, Newsletter, QR)

Ohne veröffentlichten Link sieht das Formular nur dann Chips, wenn lokal Einstellungen/Stammdaten vorhanden sind.

## Abgrenzung

Kein Ersatz für WebUntis-Absenzen/Entschuldigungen. Ideal als leichter Kanal für kleinere Schulen.

## Tool-IDs

- `krankmeldung`
- `sharepoint-liste-krankmeldung`

## Site-Key

- Settings: `ms365-krankmeldung-sp-settings-v1.siteUrl`
- Shared Kind: `krankmeldung` in `school-tool-sites`

## Smoke

- [ ] Sekretariat: Inbox, Zur Kenntnis (Actor = angemeldete Person), Archiv → Item in SPO
- [ ] KV: nur eigene Klassen; Pull/Push soft bei 403
- [ ] Eltern: Formular ohne Login, Erfolg/Fehler klar (lokal/Remote; optional SPO-Kiosk)
- [ ] Admin: Start-Pull, „Von SharePoint“, Sync-Meta unter Einstellungen
- [ ] IT: Liste einrichten (IT-Chrome), Site-Prefill
- [ ] Demo nur Admin / `?demo=1` / `?demoRole=1`; Seeds nicht nach SPO; Demo entfernen lässt echte Meldungen
