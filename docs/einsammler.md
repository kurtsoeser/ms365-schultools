# Einsammler (Eltern & Ablauf)

**Stand:** 2026-11-18  
**Status:** SPO Live Sync SoR (Familie `ablauf-kommunikation`)  
**Familie:** `ablauf-kommunikation` (Absender-Inbox + Public Magic-Link)  
**Roadmap:** [`schul-apps-roadmap-prozesse.md`](./schul-apps-roadmap-prozesse.md) · Vorlage: [`tool-upgrade-vorlage.md`](./tool-upgrade-vorlage.md) · Sync: [`spo-live-sync.md`](./spo-live-sync.md)  
**Lizenzpaket:** `ablauf` (Eltern & Ablauf)

---

## Ziel

Digitale **Einwilligungen** und **Lesebestätigungen**: Vorgang anlegen, Empfänger aus **Stammdaten-Klassen**, Fortschritt sehen, Eltern per **Magic-Link** antworten lassen.

## Rollen (kurz)

| Schicht | Wert |
|---------|------|
| Portal-Persona | Lehrkraft (Haupttool); IT für `sharepoint-liste-einsammler` |
| App-Rollen | Lehrkraft = nur eigene (`erstellerEmail`); Admin (Entra Schulleitung/IT/Verwaltung + Legacy-Mails) optional „Alle“ |
| Empfänger | Eltern ohne Microsoft-Login über Magic-Link |

## Sichtbarkeit (Lehrkraft vs. Admin)

- Jeder Einsammler speichert `erstellerEmail` / `erstellerName` (Microsoft-Anmeldung).
- **Lehrkraft** sieht in Übersicht/Liste **nur eigene** Vorgänge.
- **Admin** (E-Mails in Einstellungen, oder Stammdaten Schulleitung/Admin): Schalter „Alle Einsammler“ + Admin-Einstellungen (API, Admin-Mails, Store leeren).
- Ohne Anmeldung: keine persönliche Liste – bitte anmelden.

## Workflow (wichtig)

1. **Anmelden** (Microsoft) – Zuordnung zu Ihrer Person  
2. **Anlegen** – Titel, Modus, Typ, Frist, Text  
3. **Klassen** aus Stammdaten anhaken → Empfänger übernehmen (`Name;Klasse;Eltern-Mail`)  
4. **Aktivieren** – erzeugt persönliche Links (Unique-Token im Link)  
5. **Teilen** – Link kopieren oder „Mail vorbereiten“ (öffnet lokales Mailprogramm)  
6. **Eltern** – öffnen Link ohne Microsoft-Konto, stimmen zu / bestätigen  
7. **Übersicht** – offen / erledigt / Fortschritt  
8. **Fallback** – Status manuell setzen (Telefon etc.)

**Kein automatischer Serien-Mailversand** (noch). Später: Power Automate / Graph-Mail. Bis dahin: mailto oder Link in Teams/WhatsApp teilen.

## Funktionen

- Klassen-Chips aus Tenant-Stammdaten (+ optional Live-Hydrate über Tool-Site, local-first bei 403)  
- Empfänger aus Klassen nachladen  
- Eltern-Links je Empfänger (`einsammler-antwort.html?p=…`)  
- **Datei-/Bildabgabe:** Modus `abgabe` oder Checkbox „Upload erlaubt/Pflicht“ – Eltern laden Foto/PDF mit dem Magic-Link hoch (max. 4 MB). Speicherung lokal (IndexedDB) und optional in der Remote-API; SharePoint-Spalten `AnhangName` / `AnhangUrl` sind vorbereitet (URL-Sync folgt).  
- Antwortcode / optional Remote-API  
- SharePoint-Listen-Setup (Site über `resolveSchoolToolSiteUrl('einsammler')`)  
- **SPO Live Sync:** Köpfe + Antworten als System of Record (`einsammler-graph.js` / `einsammler-sync.js`); Pull beim Start, Push nach Absender-Schreiben; Demo-Seeds nicht pushen; 403 wipe’t Cache nicht  
- Calm Absender-UI: Action-Banner (Entwurf→Aktivieren / Links teilen), Quiet Empty, Demo nur Admin oder `?demoRole=1`

## SPO Live Sync (SoR)

| Schicht | Rolle |
|---------|--------|
| Listen `Einsammler` + `EinsammlerAntworten` | System of Record |
| Absender-App (angemeldet) | Pull/Push via Graph |
| localStorage | Cache / Offline / Demo (`localDemoOnly`, `seedTag`) |
| Public Magic-Link | **kein** Graph – Antwort via Remote-API, gleicher Browser (Inbox) oder Antwortcode; Absender übernimmt → Push nach SPO |

Module: `src/tools/einsammler/einsammler-graph.js`, `einsammler-sync.js`.  
IT: `sharepoint-liste-einsammler` legt Schema an.

### Smoke (Absender / Admin + SPO)

1. Listen-Assistent: Site + Listen anlegen.  
2. Absender: Einsammler speichern → Item in SPO-Liste `Einsammler` (`spoItemId`).  
3. Empfänger hinzufügen → Items in `EinsammlerAntworten`.  
4. Admin: „Von SharePoint laden“, „Alle Einsammler“.  
5. Public-Antwort (gleicher Browser) → Posteingang übernehmen → Antwort-Status in SPO.  
6. Demo laden: lokal, ohne SPO-Push; 403 → ruhiger Hinweis, Cache bleibt.

## Tool-IDs

- `einsammler`
- `sharepoint-liste-einsammler`

## Site-Key

- Settings: `ms365-einsammler-sp-settings-v1.siteUrl`
- Shared Kind: `einsammler` in `school-tool-sites`
