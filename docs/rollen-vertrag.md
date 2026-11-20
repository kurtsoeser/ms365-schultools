# Rollenvertrag: Persona, App-Rolle, Lizenz

**Stand:** 2026-03-09  
**Status:** verbindlich für Portal und Schul-Apps  
**Kontext:** [lehrer-portal-roadmap.md](./lehrer-portal-roadmap.md), [license-modules.md](./license-modules.md)

---

## 1. Drei Schichten (nicht vermischen)

| Schicht | Frage | Beispiele |
|--------|--------|-----------|
| **Lizenz** | Hat die *Schule* das Modul? | Pakete/`features` via License-API |
| **Portal-Persona** | Wer bin *ich* im Portal? | `it` · `lehrer` · `schueler` |
| **App-Rolle** | Was darf ich *in diesem Tool*? | SA: `admin`/`lehrer`/`schueler` · FR: `direktion`/`kv`/`schueler` |

Lizenz ohne Persona blendet für niemanden sinnvoll. Persona ohne App-Rolle zeigt die Kachel, die App entscheidet Lesen/Schreiben. UI-Checks ersetzen keine SharePoint-/Graph-Rechte.

---

## 2. Zwei Welten + IT-Kontrolle

Mentales Modell:

1. **Portal** – Alltag für Lehrkräfte und Schüler/innen  
2. **IT-Backend** – Setup, Listen, Hygiene, voller Katalog  

Nur **Schul-IT** (Global Admin, designierter IT-Kontakt, Betreiber) darf zwischen diesen Welten wechseln.

Die **Schüler-Ansicht** im Auth-Menü ist keine dritte Welt, sondern eine **Kontroll-/QA-Vorschau** („wie sieht das Portal für Schüler aus?“). Spätere Vorschauen (z. B. Schulleitung) können denselben Slot nutzen – bewusst später.

Lehrkräfte und Schüler/innen sehen **keinen** Identitäts-Umschalter. Ihre Portal-Rolle kommt aus Anmeldung, Entra-Gruppen und Stammdaten.

---

## 3. Portal-Personas

| Persona | Typische Quelle | Dashboard |
|---------|-----------------|-----------|
| Schul-IT | Entra Global Admin, License-ItContacts, Operator | Portal *oder* IT-Backend (+ optional Vorschauen) |
| Lehrkraft | Lehrer-Sammelgruppe, Planer-Rollen, Stammdaten-Match | Portal (Whitelist-Apps) |
| Schüler/in | Schüler-Sammelgruppe bzw. Fallback | Portal (Schüler-Apps) |

**Nicht** als eigene Dashboard-Persona in dieser Phase: Eltern, KV, Direktion.  
KV/Direktion sind **App-Rollen**. Eltern bleiben Magic-Link / Public-Formulare.

Code: `src/shared/dashboard-audience-resolve.js`, Cache in `dashboard-persona-session.js`.

---

## 4. App-Rollen (Muster)

Referenz: Schularbeiten- und Freistellungs-Planer (`*-entra-role.js`).

1. Rollen aus Entra-Gruppen (+ Stammdaten, z. B. KV-Mail) auflösen  
2. UI und Aktionen an aktiver Rolle ausrichten  
3. SharePoint-Item-/Listenrechte mitziehen  
4. **Demo-Umschalter** nur mit `?demoRole=1` (IT-Vorschau)  
5. **Mehrfachrollen-Switcher** nur bei echten Mehrfachrollen (Policy: `src/shared/app-role-switcher-policy.js`)  
6. Lokale Demo-Chips („Rolle (Demo)“ in localStorage) sind produktiv abzuschaffen, sobald Gruppen greifen  

---

## 5. Dashboard-Sichtbarkeit

Pipeline:

1. Lizenz-Feature vorhanden?  
2. Layout: IT-Backend (`full`) vs. Portal (`preview-*` / `apps-*`)  
3. Portal-fähig (`DASHBOARD_LEHRER_SCHUELER_APPS`) – Alltags-Apps, kein IT-Setup  
4. Effektive Audience (Default-Regel + IT-Matrix `ms365-dashboard-tool-access-v1`)  

IT-Setup-Tools bleiben `audience: it` und außerhalb der Portal-Whitelist. Schul-Apps sind portal-fähig und haben passende Audience (`lehrer` / `schueler` / `all`). Die IT-Matrix unter Werkzeug-Zugriff steuert, welche portal-fähigen Apps Lehrkräfte bzw. Schüler sehen.

**Schul-Portal-Standard (P0.2):** Der IT-Leiter definiert angeheftete Home-Apps und Sidebar-Reihenfolge pro Persona (`ms365-dashboard-portal-defaults-v1`, UI „Portal gestalten“). Persönliche Favoriten/Canvas bleiben User-seitig und werden nicht mit dem Schul-Backup vermischt. Schichten: Lizenz → portal-fähig → IT-Schul-Standard → optional persönlich.

---

## 6. Wortlaut

| Begriff | Bedeutung |
|---------|-----------|
| Schul-IT | Portal-Persona / Shell-Recht für Backend |
| Direktion / Admin-Inbox | App-Rolle (Genehmigen, alle Anträge) – nicht gleich Schul-IT |
| Klassenvorstand (KV) | App-Rolle aus Stammdaten/Gruppen |
| Vorschau | Nur IT: Portal-Ansicht einer anderen Zielgruppe testen |

---

## 7. Wellen (Umsetzung)

| Welle | Inhalt |
|-------|--------|
| 0 | Dieser Vertrag, Shared Switcher-Policy, Katalog-Audit, Persona im Auth-Menü – erledigt |
| 1 | Feinschliff SA/FR/Termine/Teams-Antrag – erledigt |
| 2 | Demo-Chips → Entra (Aktivitäten, Lehrer-Freistellung, Projektwochen, Wahlfach) – erledigt |
| 3 | Ablauf-Apps (Einsammler, Elternbrief, Krankmeldung, Kollegium) – erledigt |
| 4 | Betrieb-Apps (Aufgaben, Technik, Raumplan, Ausgaben) – erledigt |

### Welle 3 (Ablauf)

Shared Helper: `src/shared/ablauf-admin-entra.js` (Schulleitung / Verwaltung / SA-Admin / Schul-IT).

| App | Admin | Audience / Scope |
|-----|-------|------------------|
| Einsammler / Elternbrief | Ablauf-Admin (+ Legacy-Mails) | Owner vs. Admin `showAll` |
| Krankmeldung | Ablauf-Admin (Sekretariat) | KV nur eigene Klassen; Eltern weiter Public/Token |
| Kollegium | Ablauf-Admin | `alle_lehrkraefte` nur Lehrer-Gruppe/Stammdaten; KV-/Abteilungs-Zielgruppen |

### Welle 4 (Betrieb)

Shared Helper: `src/shared/betrieb-admin-entra.js` (nutzt Ablauf-Admin).

| App | Lehrkraft / Schüler | Admin (Betrieb) |
|-----|---------------------|-----------------|
| Aufgaben | anlegen; Scope mir/von_mir (Lehrer) | `showAll` + Clear/Demo |
| Technik | melden; nur eigene (Lehrer **und** Schüler) | Inbox + Status/Löschen |
| Raumplan | lesen/suchen + Technik-Link (Lehrer) | Import, CRUD, Plan-Upload |
| Ausgaben | „meine“ Vorgänge (`empfaengerEmail`) | Inventar/Ausgeben/Retour/Settings; Audience `lehrer` |

---

## 8. Diagnose, wenn die falsche Ansicht offen ist

Die App-Rolle wird beim Öffnen berechnet. Quelle ist das angemeldete Konto plus die Einrichtung **dieses** Tools (Genehmiger, Personen, Entra-Gruppe, Schulleitung in den Stammdaten, globaler Administrator). Die IT-Bibliothek speichert diese Einrichtung (`permissions-*.json`) und lädt sie in den Browser. Sie vergibt die Rolle nicht.

Die Portal-Persona im Kontomenü ist eine andere Schicht. Schulleitung bleibt dort oft „Lehrkraft“.

Anzeige neben dem Titel: `App-Rolle · <Quelle>`.

| Anzeige | Bedeutung | Nächster Schritt |
|---------|-----------|------------------|
| Demo-Direktor, Genehmiger (Setup), Schulleitung (Stammdaten), Entra-Gruppe | Konto wurde gelesen, Quelle ist klar | Listenrechte und eigene/alle Zeilen prüfen |
| `default` oder „Konto wird gelesen“ | Rolle wurde ohne E-Mail festgelegt | Anmeldung abwarten, nichts in SharePoint ändern |
| Direktion-Gruppe leer, Person ist eingetragen | Zugriff läuft über die Person | Gruppe nicht als Fehler lesen |

Eine Rolle darf erst gespeichert werden, wenn die E-Mail des Kontos gelesen ist. Ein leerer früher Lauf darf einen späteren Lauf mit Konto nicht überschreiben.
