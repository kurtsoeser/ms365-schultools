# Smoke-Checkliste: Schuljahreswechsel

Manuell nach Fixes aus Analyse 01 (Korrektheit). Haken setzen, wenn bestanden.

## Vorbereitung

- [ ] Browser-Backup exportiert
- [ ] Demo-Daten oder Testdaten mit mind. zwei Schuljahren
- [ ] Angemeldet mit Graph-Rechten (Lesen für Match, Schreiben nur bewusst)

## Schuljahr-Label (K5)

- [ ] Systemdatum bzw. Anzeige: Jan–Aug zeigt Vorjahr als Start (z. B. März → `YYYY-1/YY`)
- [ ] Ab 1.9. neues Schuljahr in Stammdaten / Org-Assistent-Vorschlag

## Match (K1)

- [ ] Tenant mit `1A` und `11A`: Auto-Match verknüpft `1A` nicht mit `11A`
- [ ] Mehrdeutige Namen (z. B. „Max“) → kein Auto-User-Match

## Sync / Truncation (K2/K3)

- [ ] Schüler mit Nummern-UPN + Alias in Stammliste: Diff zeigt „beide“, kein fälschliches Entfernen
- [ ] Bei gekürzter Mitgliederliste: Sync bricht ab, Toast/Log warnt, keine Leaves

## State / Multi-Tab (K4)

- [ ] Org-Assistent speichert Playbook → Gruppenverwaltung (Struktur) neu laden: Playbook noch da
- [ ] Optional: zweiter Tab ändert Struktur → erster Tab aktualisiert nach Fokus/Storage

## Merge / Umbenennen

- [ ] Klassen-Merge ohne Graph-Match: Warnung, nur lokal nach Bestätigung
- [ ] Klassen-Umbenennen: Nickname mit Umlauten (`ä`→`ae`)

## Cleanup

- [ ] Keine unerwarteten Gruppen-Deletes
- [ ] Backup wieder importierbar
