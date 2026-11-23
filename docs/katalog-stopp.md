# Katalog-Stopp (a11)

**Stand:** 21.11.2026  
**Entscheidung:** Bis die Kern-Apps und der Schuljahreswechsel herausragen, keine neuen Werkzeuge. Bestehende Seiten sind **Kern**, **Pflege** oder **Archiv**.

## Bedeutung

| Stufe | Bedeutung |
|---|---|
| **Kern** | Alltags-/Verkaufsfokus (Portal-Planer, Elternkanal, Kursteams, Jahreswechsel, Organigramm) |
| **Pflege** | Bleibt nutzbar (IT-Setup, Listen, Hygiene) – kein Feature-Ausbau |
| **Archiv** | Keine Weiterentwicklung; Kandidaten zum Ausblenden |

## Quelle im Code

`src/shared/tool-lifecycle-catalog.js` – Overrides + Heuristik aus `DASHBOARD_LEHRER_SCHUELER_APPS`.

Zählen: `lifecycleCounts()` (Test deckt ab, dass jedes `DASHBOARD_TOOL_LABELS`-Werkzeug klassifiziert ist).

## Anti-Ziele

- Kein neues Werkzeug, solange Jahreswechsel und die drei Kern-Apps nicht herausragend sind.
- Jedes neue ersetzt ein altes (Archiv → raus).
