# Teams-Antrag

**Stand:** 2026-11-07  
**Status:** Spezifikation + MVP (local-first)  
**Lizenz:** `teams-antrag` in Paketen **`core`** und **`ablauf`** (Union – wer eines der beiden hat, sieht die Kachel)  
**Verwandt:** [`gruppenerstellung`](../tools/gruppenerstellung.html) (Einschränkung), [`weitere-teams-gruppen`](../tools/weitere-teams-gruppen.html) (IT-Wizard), [`diplomarbeiten`](./diplomarbeiten.md)

---

## Ziel

Lehrkräfte **und Schüler:innen** beantragen ein neues Microsoft Team über die **schulische Entra-ID**. Nur die IT (bzw. ein konfigurierbarer Freigeber-Kreis) prüft, setzt die Nomenklatur und legt das Team **automatisch per Graph** an. So bleibt die Tenant-Hygiene sauber, obwohl die Gruppenerstellung für normale Nutzer:innen gesperrt sein kann.

## Nicht im Scope

- Kursteams / Education-Class (eigenes Tool)
- Diplomarbeiten / VWA (eigenes Tool – Antragstyp verweist dorthin)
- Öffentliches Formular ohne Login
- Automatischer Serien-Mailversand (optional später Power Automate)

## Rollen

| Rolle | Kann |
|-------|------|
| Schüler:in / Lehrkraft (Entra) | Antrag stellen, eigene Anträge + Status sehen |
| Freigeber (IT) | Warteschlange, ablehnen, zurückstellen, Name/Alias finalisieren, genehmigen → Graph-Anlage |
| Admin | Freigeber-Gruppe zuordnen, Extra-Mails, Nomenklatur, Demo/Reset |

### IT-Übersichten

| Ansicht | Inhalt |
|---------|--------|
| **Offen (IT)** | Nur Warteschlange (beantragt / in Prüfung / fehlgeschlagen) |
| **Alle Anträge** | Gesamtübersicht aller Status – Suche + Statusfilter, Detail mit Freigabe-Aktionen |

### Wer darf freigeben?

Standard: Mitglieder der **selben reinen Sicherheitsgruppe**, die unter [`Wer darf Teams anlegen?`](../tools/gruppenerstellung.html) (`Group.Unified` / `GroupCreationAllowedGroupId`) hinterlegt ist – **nicht** mail-aktiviert.

In den Tool-Einstellungen:

1. **Aus „Wer darf Teams anlegen?“ übernehmen** (lokaler Stand) oder **aus Mandant (Group.Unified) lesen**
2. optional Gruppe suchen / GUID setzen
3. optional einzelne Extra-Mails (Ausnahmen)

Zusätzlich: Schul-IT laut Lizenzkontakt. Für Bulk-Anlage: Link zu [`Weitere Teams & Gruppen`](../tools/weitere-teams-gruppen.html).

Ohne Anmeldung: kein persönlicher Antrag – bitte Microsoft-Konto der Schule.

## Statusfluss

```text
entwurf → beantragt → in_pruefung → genehmigt → angelegt
                              ↘ abgelehnt | zurueckgestellt
genehmigt → (Fehler) anlage_fehlgeschlagen → Retry → angelegt
```

## Felder (Antrag)

| Intern | Bedeutung |
|--------|-----------|
| `Title` | Kurzname / Arbeitsitlel |
| `AntragId` | stabile ID (`ta-…`) |
| `Zweck` | Begründung (Pflicht) |
| `Art` | `projekt` \| `ag` \| `fachgruppe` \| `klasse` \| `diplom_vwa` \| `sonstiges` |
| `Namensvorschlag` | Wunschname (nicht final) |
| `BesitzerMails` | gewünschte Besitzer (Freitext, Zeilen/Semikolon) |
| `MitgliederMails` | gewünschte Mitglieder |
| `Schuljahr` | z. B. 2026/27 |
| `AntragstellerEmail` / `AntragstellerName` | aus MSAL |
| `Status` | siehe oben |
| `FinalDisplayName` / `FinalMailNickname` | von IT gesetzt |
| `TeamId` / `TeamWebUrl` | nach Anlage |
| `Ablehnungsgrund` / `NotizIt` | Freitext IT |
| `Quelle` | `ui` \| `demo` |

Bei Art `diplom_vwa`: UI-Hinweis + Link zum Diplom-/VWA-Tool – **kein** generisches Team anlegen.

## IT bei Genehmigung

1. Finalen Anzeigenamen + Mail-Nickname setzen (Preview)
2. Besitzer (mind. Antragsteller oder IT) festlegen
3. Genehmigen → Graph: `createUnifiedGroup` + `provisionTeamForGroup`
4. Antragsteller:in als Mitglied; Status `angelegt` + Team-Link

## Architektur

```text
MSAL (Schüler/Lehrkraft/IT)
    → Tool-UI (localStorage MVP)
    → optional SharePoint-Liste TeamAntraege (SoR später)
    → Graph (nur Freigeber) Unified Group + Team
```

MVP speichert in `localStorage` (`ms365-teams-antrag-v1`), analog Einsammler/Ausgaben. SharePoint-Schema ist vorbereitet.

## Tool-IDs

- `teams-antrag`
- `sharepoint-liste-teams-antrag`

## Dashboard

- Cluster: `gruppen` (neben Weitere Teams) und sichtbar unter Ablauf-Audience
- Audience: `all` (Lehrkraft + Schüler:in + IT; Freigabe-UI nur für Freigeber)

## Dateien

- `tools/teams-antrag.html`
- `src/tools/teams-antrag/*`
- `tools/sharepoint-liste-teams-antrag.html` (Setup)
