# Schuldaten-Karte (Schulgraph)

Interaktive **Beziehungsvisualisierung** der lokalen Schul-Stammdaten: zoombarer SVG-Graph mit Layern, Fokus und Detail-Panel.

## Ziele

- Querschnitt über Listen und Relationen, die heute in Einzel-Tools stecken (Stammdaten, Kursteams/Belegung, Einrichtung, Planer).
- **Erkunden**, nicht bearbeiten – Pflege bleibt in `tenant.html`, Kursteams, SharePoint-Listen usw.
- Datenschutz: bei vielen Schüler:innen wird der Personen-Layer standardmäßig ausgeschaltet; Filter nach Klasse.

## Architektur

| Modul | Rolle |
|--------|--------|
| `schulgraph-schema.js` | Knoten-/Kantentypen, Legende |
| `schulgraph-aspects.js` | Aspekte, Presets, Normalisierung |
| `schulgraph-logic.js` | `buildSchulGraph()` – reine Daten → Graph |
| `schulgraph-ui.js` | Shell, SVG, Sidebar |
| `schulgraph.js` | Wiring, Pan/Zoom, Persistenz der Layer |
| `tools/schulgraph.html` | Tool-Seite |

## Datenquellen

1. **`ms365TenantSettingsLoad()`** – Klassen, Fächer, ARGE, Lehrkräfte (+ ggf. Schüler in Core)
2. **`ms365AppDataV2.getYearBucket()`** – Schüler, Erziehungsberechtigte, `unterrichtsbelegung`
3. **`setup.catalogLinks`** – Verknüpfung Stammdaten ↔ M365-Gruppen (Einrichtungsassistent)

## Kanten (Auswahl)

| Kante | Bedeutung |
|--------|-----------|
| `belongs_to` | Klasse/Fach/ARGE → Schul-Knoten |
| `kv_of` | Klassenvorstand → Klasse |
| `subject_in_arge` | Fach → ARGE |
| `teaches` / `class_subject` | Unterrichtsbelegung (Lehrkraft, Klasse, Fach) |
| `in_class` / `guardian_of` | Schüler, Eltern |
| `m365_link` | Stammdaten → Gruppe |

## Aspekte (was der Graph zeigt)

Presets sind Schnelleinstiege; darunter können **Beziehungs-Aspekte** einzeln ein-/ausgeschaltet werden (Speicherung: `localStorage` `ms365-schulgraph-options-v2`).

| Aspekt | Inhalt |
|--------|--------|
| Schulorganisation | Klassen, Fächer, ARGE am Schulzentrum |
| Klassenvorstand | Lehrkraft → Klasse (KV) |
| Fächer & ARGE | Fach → ARGE/Fachgruppe |
| Unterrichtsbelegung | Lehrkraft ↔ Klasse ↔ Fach (Kursteams) |
| Schüler:innen | Schüler:in → Klasse |
| Erziehungsberechtigte | Eltern → Schüler:in (nur mit Schüler:innen) |
| Microsoft 365 | `catalogLinks` → Entra-Gruppen |

**Presets:** Gesamtüberblick · Unterricht · Fächer & ARGE · Klassen & KV · Familie & Klasse · Microsoft 365 · Frei kombinieren

## Bedienung (Kernidee)

**Klick auf einen Knoten** lädt die **Nachbarschaft** und zeigt sie mit Verbindungen im Graph:

| Angeklickt | Was passiert |
|------------|----------------|
| **Lehrkraft / Schüler:in / Eltern** (mit E-Mail) | Gruppen aus **Microsoft 365** (`memberOf`) – Linien „Mitglied in“ |
| **M365-Gruppe** | **Mitglieder** der Gruppe aus Entra (Knoten „M365-Benutzer“) |
| **Klasse** | Mitglieder der **verknüpften Klassen-Gruppe** (Einrichtung `catalogLinks`); sonst Schüler:innen aus **lokalen Stammdaten** |

- **Anmeldung** oben (MSAL) erforderlich für Live-Daten aus dem Tenant.
- **Nachbarschaft schließen** entfernt die zuletzt geladenen Zusatz-Knoten.
- **Doppelklick** = Fokus (nur Nachbarschaft im Graph sichtbar).
- **Mausrad** = Zoom, **Ziehen** = Schwenken.

Die **Aspekte/Presets** steuern weiterhin, welche Stammdaten-Struktur im Hintergrund-Graph sichtbar ist.

## Roadmap (Vorschläge)

- [ ] Aspekt **Schularbeiten** (Termine aus SharePoint-Planer)
- [ ] Aspekt **Projektwochen / Aktivitäten**
- [ ] Export PNG/SVG, Deep-Link `?focus=class:3LA`
- [ ] Optional D3-Force-Layout für große Graphen
- [ ] Suche / Autocomplete über Knotenlabels
- [ ] Rechte: Personen-Layer nur für Verwaltung

## Tests

```bash
npm test -- tests/schulgraph-logic.test.mjs
```
