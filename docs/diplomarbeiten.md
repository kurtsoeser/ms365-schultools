# Diplomarbeiten & vorwissenschaftliche Arbeiten (VWA)

**Stand:** 2026-11-07  
**Status:** Spezifikation + Umbau Option C  
**Lizenzpaket:** `unterricht` (`toolId`: `diplomarbeiten`)  
**Verwandt:** [`teams-antrag`](./teams-antrag.md), sRDP-Profile (`arbeitLabel`), PA `pa-diplom-ordner`

---

## Ziel

Abschlussarbeiten **koordiniert** in Microsoft Teams verwalten – nicht jeder legt eigene Teams an.

**Standard (empfohlen):** ein **Dach-Team pro Maturajahrgang**, Kanal „Allgemein“ für alle Infos, **private Kanäle je Arbeit** (Betreuung + 3–5 Schüler:innen).

**Legacy:** bisheriges Modell – eine M365-Gruppe/Team pro Arbeit (`dipl-{jahr}-…`).

## Schulform-Bezeichnung

Ein Tool, dynamische Labels:

| Schulform | Bezeichnung in der UI | Naming-Präfix (Kanäle / Legacy) |
|-----------|------------------------|----------------------------------|
| AHS | Vorwissenschaftliche Arbeit (VWA) / ABA | `vwa-` |
| BHS (HTL, HAK, HLW, …) | Diplomarbeit | `da-` / `dipl-` |

Einstellung **Abschlussarbeit-Typ** (`diplom` \| `vwa`) im Setup; Kachel-/Hilfetexte folgen dem Typ.

## Setup-Wizard (Option C)

1. Abschlussarbeit-Typ wählen (`diplom` / `vwa`)
2. Betriebsmodus:
   - **Standard:** Dach-Team + private Kanäle
   - **Legacy:** eigene Teams/Gruppen pro Arbeit
3. Bei Standard: Maturajahrgang anlegen → Team erzeugen oder bestehende Team-ID verknüpfen

Moduswechsel jederzeit möglich; Legacy-Gruppen bleiben auffindbar.

## Standard-Modus

```text
Maturajahrgang 2026/27
  └── Team „Diplomarbeiten 2026/27“ (bzw. „VWA 2026/27“)
        ├── Allgemein
        └── Private Kanäle je Arbeit
Maturajahrgang 2027/28
  └── neues Dach-Team …
```

### Daten (Arbeit)

| Feld | Bedeutung |
|------|-----------|
| Titel / Thema | Arbeitstitel |
| Betreuer | E-Mail / Name |
| Schüler:innen | 1–n UPNs/Mails |
| Status | `geplant` → `aktiv` → `abgeschlossen` → `archiviert` |
| TeamId / ChannelId | Graph-Referenzen |
| KanalName | z. B. `DA – Müller / KI` |

### Aktionen / UI

- Jahrgangsliste → **Jahrgang öffnen** (verschachtelt)
- Dach-Team: **neues Team anlegen** (M365-Gruppe + Teams-Provisionierung) **oder** bestehendes Team/Gruppe **suchen und verbinden** (fehlt Teams-Funktionalität, wird nachgezogen)
- Arbeit: Betreuung & Schüler:innen per **Entra-Suche** wählen → speichern → **privaten Kanal** anlegen
- Abschließen → Status; Kanal archivieren/löschen in Teams (IT)

## Legacy-Modus

Unverändert nutzbar:

- Anzeigename `Diplomarbeit {Jahr} – Thema` (bzw. VWA-Variante)
- Alias `dipl-{jahr}-{slug}` / `vwa-{jahr}-{slug}`
- Optional nur Gruppe ohne Team

## Abgrenzung Teams-Antrag

Anträge mit Art „Diplom/VWA“ werden im Teams-Antrag **umgeleitet** – Abschlussarbeiten gehören in dieses Tool.

## Tool-IDs

- `diplomarbeiten`
- `pa-diplom-ordner` (Power-Automate-Rezept, unverändert)

## Dateien

- `tools/diplomarbeiten.html`
- `src/tools/diplomarbeiten/*`
