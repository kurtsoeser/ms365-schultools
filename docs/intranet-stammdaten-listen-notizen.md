# Intranet-Stammdaten-Listen – Abgrenzung & Berechtigungsmodell

Stand: Produktnotiz für spätere Ausbaustufen.

## Was im Schulregister (Stammdaten) liegt

- **SharePoint-Website** für Intranet-Listen (`intranetSiteUrl`)
- **Listenname je Typ** (`setup.intranetListTitles`) – Standards, wenn leer

**Nicht** im Stammdaten-Tab des Registers (vorerst):

- Block **„Abgleich & Berechtigungen“** (Immer neue Listen, Verwaiste Zeilen, Berechtigungen überspringen, Entra-Gruppen-Picker)
- Diese Optionen bleiben im Werkzeug `tools/sharepoint-liste-stammdaten.html` bzw. werden später an einer zentralen Stelle (IT/Setup) gebündelt – nicht im Schulregister-Tab (der Tab „Intranet-Listen“ wurde entfernt).

## Abgleich & Berechtigungen – Wirkung heute

- **Abgleich** (Sync-Modus, verwaiste Zeilen): wird von der Intranet-Einbettung und den Schnell-Sync-Buttons im Register genutzt (`stammdaten-intranet-listen-ui.js`, `sharepoint-liste-stammdaten.js`).
- **Berechtigungen**: Entra-Gruppen → SharePoint-Listenrollen nach Profil (`stammdaten-liste-permissions.js`, analog Schularbeiten-Planer). Konfiguration aktuell **localStorage** (`ms365-stammdaten-listen-perms-v1`), Gruppen-Picker in der UI.

## Merkhilfe: Wer ist „Lehrkraft“ / „Schüler“ für Berechtigungen?

Für die spätere zentrale Konfiguration gilt inhaltlich:

| Rolle in den Berechtigungen | Bedeutung |
|----------------------------|-----------|
| **Lehrkräfte-Gruppe** | Die **Microsoft-365-Gruppe der Lehrkräfte** aus dem Register (Lehrerliste ↔ Sammelgruppe / Entra), nicht eine beliebige Gruppe. Wer in der **Lehrerliste** geführt wird, gehört fachlich zu dieser Lehrkraft-Welt; Listen-Berechtigungen „Lehrer“ beziehen sich auf diese Gruppe. |
| **Schüler-Sammelgruppe** | Die **zugehörige Microsoft-365-Gruppe der Schüler:innen** aus dem Register (Schülerliste ↔ Sammelgruppe). Berechtigungsprofil „Schüler“ auf Listen (z. B. Lesen bei Fächern/Klassen) bezieht sich auf diese Gruppe – nicht auf die Gesamtschülerliste als Personenfeld-Inhalt allein. |

Personenfelder in Listen (z. B. Klasse → Schüler als M365-Personen, Lehrer → Lehrkraft) sollen **konsistent** zu den Stammdaten und den genannten Sammelgruppen sein; die SharePoint-Rollen kommen über die **Entra-Gruppen-Picker** („Verwaltung“, „Lehrkräfte“, „Schüler (Sammelgruppe)“).

## Offen (bewusst zurückgestellt)

- Abgleich- und Berechtigungs-Defaults **nicht** in `app-data-v2` Setup migrieren, bis der Zielort (IT-Panel / ein Setup-Wizard) feststeht.
- Optional: Gruppen aus Register (`schuelerGroupId`, Lehrer-Sammelgruppe) automatisch in die Berechtigungs-Picker vorbelegen.
