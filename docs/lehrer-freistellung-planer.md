# Freistellungen Lehrkräfte

Eigenes Werkzeug neben dem **Schüler-Freistellungs-Planer** (`freistellung-planer`).

## Ziel

- Lehrkräfte stellen Freistellungs-/Abwesenheitsanträge (Fortbildung, Arzt, …).
- **Nur die Direktion** genehmigt (ein Approvals-Schritt in Power Automate).
- Übersicht, **Kalender** und **iCal-Export** für Sekretariat / freigegebenen Outlook-Kalender.

## Technik (Stand)

| Bereich | Pfad |
|--------|------|
| UI | `tools/lehrer-freistellung-planer.html` |
| Modul | `src/tools/lehrer-freistellung-planer/` |
| IT-Setup | `tools/lehrer-freistellung-setup.html` |
| Setup (Browser) | `localStorage` `ms365-lfr-setup-v1`, `ms365-lfr-perms-v1` |
| SharePoint-Liste | Standardtitel `Lehrer-Freistellungen` (Schema in `lfr-schema.js`) |

Ohne konfigurierte Site läuft ein **Demo-Modus** (lokale Beispieldaten).

## Power Automate

Vorlage: `assets/power-automate/lehrer-freistellung/` (Flow-ID `b8d4e2f1-…`). Paketbau im IT-Setup Schritt 5.

1. Trigger: neues Listenelement (Status `Ausstehend`).
2. **Eine** Genehmigung an die Direktion (E-Mail aus Setup Schritt 2).
3. Ergebnis → Status, `BemerkungDirektion`, Audit-Felder `GenehmigtVonDirektion` / `GenehmigtAmDirektion` bzw. `AbgelehntVon` / `AbgelehntAm`.
4. Status-Mail an `LehrerEmail`.

Vorlage neu erzeugen: `node scripts/build-lfr-pa-template.mjs`

## Kalender

- **iCal:** `lfr-export.js` im Planer (Ansicht Kalender-Export).
- **Graph:** `lfr-calendar-sync.js` – genehmigte Termine in Kalender eines freigegebenen Postfachs (`outlookCalendarUser` / optional `outlookCalendarId` im IT-Setup Schritt 4). Event-IDs werden lokal gemappt (`ms365-lfr-outlook-event-v1`).
