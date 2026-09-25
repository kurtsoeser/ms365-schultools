# Prompt: Lehrplan → Kursteam-Kanal-Vorlagen (JSON)

Kopiere den gesamten Block unter **PROMPT** in eine KI deiner Wahl. Lade zusätzlich den Lehrplan (PDF/Text) hoch. Die KI soll **nur gültiges JSON** zurückgeben, das du in MS365-Schul-Tools unter **Kursteam-Vorlagen → Import** einspielen kannst.

Referenzdatei im Repo: `assets/kursteam-templates/vorlagen-import.example.json`

---

## Begriffe (österreichisches Schulsystem)

| Feld | Bedeutung |
|------|-----------|
| **schulstufe** | Absolute Schulstufe 1–13 (bzw. höher bei Sonderfällen). Unabhängig von der Klassennummer der Schulform. |
| **semester** | `SJ` = ganzes Schuljahr (WS+SS), `WS` = Wintersemester, `SS` = Sommersemester. Optional. |
| **schoolForm** | Schulform (AHS, HAK, HAKB, …). |
| **subjectCode** | Fachkürzel (MAM, M, D, …). |

**Klassennummer → Schulstufe (Faustregel):**

- **AHS / Mittelschule:** 1. Klasse = Stufe **5** … 8. Klasse = Stufe **12**
- **BHS** (HAK, HAS, HTL, HLW, BAfEP, …): 1. Klasse = Stufe **9** … 5. Klasse = Stufe **13**

Beispiel: HAK 3. Klasse → `schulstufe: "11"`. AHS 5. Klasse → `schulstufe: "9"`.

**HAKB/HAK/HAS-Lehrplan-„Modul“ (Halbjahr, z. B. MAM):** ungerade = WS, gerade = SS; Stufe = 8 + ⌈Modul/2⌉

| Modul | Schulstufe | Semester |
|-------|------------|----------|
| 3 | 10 | WS |
| 4 | 10 | SS |
| 5 | 11 | WS |
| 6 | 11 | SS |
| 7 | 12 | WS |
| 8 | 12 | SS |

Alternative im JSON: statt `schulstufe` darf `klasse` (Klassennummer) **oder** bei HAKB/HAK/HAS `module` (Lehrplan-Modul) gesetzt werden – der Import rechnet Stufe/Semester aus.

---

## PROMPT

```text
Du erstellst Kanal-Vorlagen für Microsoft Teams Kursteams (Schulverwaltung Österreich).

## Ziel
Aus dem angehängten Lehrplan (bzw. den hochgeladenen Dokumenten) erzeugst du eine JSON-Datei mit mehreren Vorlagen. Jede Vorlage beschreibt die empfohlenen STANDARD-Kanäle für ein Unterrichtsteam zu einer Kombination aus:

- Schulform (schoolForm)
- Fach (subjectCode)
- Österreichische Schulstufe (schulstufe) – absolute Stufe, nicht die Klassennummer der Schulform
- Optional Semester (semester): SJ = ganzes Schuljahr (WS+SS), WS, SS

## Ausgabeformat (streng einhalten)
Gib AUSSCHLIESSLICH ein einziges JSON-Objekt aus – kein Markdown, keine Erklärung davor/danach.

Schema:
{
  "kind": "ms365-kursteam-templates",
  "version": 3,
  "exportedAt": "<ISO-8601-Zeitstempel>",
  "templates": [
    {
      "name": "string – sprechender Anzeigename",
      "schoolForm": "string – z. B. AHS | Mittelschule | HAK | HAKB | HAS | HTL | HLW | BAfEP | PTS | Berufsschule",
      "subjectCode": "string – Fachkürzel GROSS, z. B. MAM, M, D, E, GWK",
      "schulstufe": "string – absolute österr. Schulstufe, z. B. \"5\", \"9\", \"11\"",
      "semester": "string – optional: \"SJ\" | \"WS\" | \"SS\"",
      "description": "string – 1–2 Sätze: Lehrplanbezug / Inhaltsschwerpunkt",
      "channels": [
        "01 - Thema …",
        "02 - Thema …"
      ]
    }
  ]
}

## Regeln für channels
1. NICHT eintragen: „Allgemein“, „General“, „00-Allgemein“ (existiert im Team bereits).
2. Nur Standard-Kanäle (keine privaten Kanäle).
3. Pro Vorlage sinnvolle Lern-/Themenblöcke als Kanäle (typisch 5–12, max. 15).
4. Kanalnamen möglichst ≤ 50 Zeichen.
5. Für stabile Sortierung in Teams: zweistellige Nummer + Bindestrich, z. B. „01 - …“, „02 - …“.
6. Optional Emojis nur wenn im Lehrplan-/Schulstil üblich und die Länge erlaubt.
7. Keine Duplikate innerhalb derselben Vorlage (case-insensitive).
8. Keine leeren strings in channels.

## Regeln für Metadaten
1. schoolForm muss zur Zielschule passen. Wenn der Lehrplan mehrere Schulformen abdeckt, lege getrennte Vorlagen an.
2. subjectCode: bekannte Kürzel der Schule verwenden, falls im Dokument genannt; sonst sinnvolle österreichische Kürzel.
3. schulstufe: IMMER die absolute österreichische Schulstufe (nicht die Klassennummer der Schulform).
   Umrechnung Klasse: AHS/MS 1.=5 … 8.=12; BHS (HAK/HTL/…) 1.=9 … 5.=13.
   HAKB/HAK/HAS-Lehrplan-Module (Halbjahre): Modul 3=Stufe 10 WS, 4=10 SS, 5=11 WS, 6=11 SS, 7=12 WS, 8=12 SS.
   Wenn der Lehrplan nur Modulnummern nennt: in schulstufe+semester umrechnen; Modulnummer darf in name/description stehen.
4. semester: Bei Halbjahres-Lehrplänen (HAKB-Module) WS bzw. SS setzen. Ganzjährig → "SJ". Sonst weglassen oder "SJ".
5. name: klar und eindeutig, idealerweise „{Schulform} {Fach} Schulstufe {n} {WS|SS} – Kurzthema“.
6. description: knapp, fachlich, ohne Marketing-Floskeln; Modul-/Klassennummer der Schulform darf hier stehen.
7. id-Felder NICHT setzen (werden beim Import erzeugt).
8. Das veraltete Feld „module“ möglichst NICHT verwenden (außer bekannte HAKB-Modulnummern ohne schulstufe).

## Abdeckung
- Erzeuge Vorlagen für alle im Lehrplan klar abgrenzbaren Schulstufen (ggf. Semester) des gewünschten Fachs.
- Wenn der Nutzer eine Schulform/Fächerliste vorgibt, halte dich strikt daran.
- Wenn Informationen fehlen: sensible Annahmen in description erwähnen, aber trotzdem gültiges JSON liefern.
- Keine erfundenen Rechtsgrundlagen zitieren.

## Qualitätscheck vor Ausgabe
- JSON parsebar
- kind === "ms365-kursteam-templates"
- version === 3
- templates ist nicht-leeres Array
- jedes Element hat name, schoolForm, subjectCode, schulstufe, description, channels[]
- channels ohne Allgemein/General
```

---

## Zusatzprompt (optional, anhängen)

Wenn du die Ausgabe eingrenzen willst:

```text
Zusatzbedingungen:
- Schulform(en): <z. B. nur HAKB>
- Fächer: <z. B. nur MAM und WIN>
- Schulstufen: <z. B. 10 bis 12 (= HAKB Module 3–8)>
- Semester: <z. B. nur SJ, oder WS/SS getrennt>
- Sprache der Kanalnamen: Deutsch
- Nummerierungsstil: "01 - Titel" (mit Leerzeichen um den Bindestrich)
```

---

## Nach dem Import in der App

1. Kursteam-Vorlagen öffnen → **Import**
2. Filter: **Schulform → Fach → Schulstufe → Semester**
3. Gruppierung nach Bedarf: Schulform / Fach / Schulstufe / Semester
4. Vorlage prüfen, ggf. Kanalnamen kürzen, dann auf ein Team **Anwenden**
