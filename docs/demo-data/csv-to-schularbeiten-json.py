#!/usr/bin/env python3
"""CSV Schularbeiten-Termine → Planer-Import-JSON (einmalig / Wartung)."""
import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path

FACH_NAMES = {
    "D": "Deutsch",
    "BWL": "Betriebswirtschaftslehre",
    "ENWS": "Englisch",
    "FRWS": "Französisch",
    "SPWS": "Spanisch",
    "ITWS": "Italienisch",
    "WINF": "Wirtschaftsinformatik",
    "MAM": "Mathematik",
    "OMAI": "Angewandte Mathematik",
    "UNCO": "Unternehmensrecht und -controlling",
    "CASES": "Case Studies",
}

FACH_COLORS = [
    "#6366f1", "#0ea5e9", "#10b981", "#f59e0b", "#ef4444",
    "#8b5cf6", "#e11d48", "#64748b", "#14b8a6", "#f97316", "#78716c",
]


def parse_dt(s: str) -> datetime | None:
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"Datum nicht parsebar: {s!r}")


def dauer_aus_beginn_ende(beginn: datetime, ende: datetime | None, eh: int) -> int:
    if ende and ende > beginn:
        mins = int((ende - beginn).total_seconds() // 60)
        if mins > 0:
            return mins
    return max(50, eh * 50)


def semester_for(d: datetime) -> str:
    # SJ 2026/27: WS bis Ende Jänner, ab Februar SS
    if d.year == 2027 and d.month >= 2:
        return "SS"
    return "WS"


def infer_fach(row: dict) -> str:
    f = (row.get("Fach") or "").strip()
    if f:
        return f
    titel = (row.get("Titel") or "").strip()
    if re.match(r"^BW\b", titel, re.I):
        return "BWL"
    return ""


def parse_json_array(s: str) -> list:
    s = (s or "").strip()
    if not s:
        return []
    return json.loads(s)


def jg_year(jg: str) -> int:
    m = re.search(r"(\d+)", jg or "")
    return int(m.group(1)) if m else 0


def main() -> int:
    csv_path = Path(
        sys.argv[1]
        if len(sys.argv) > 1
        else r"c:\Users\KurtSöser\Downloads\Schularbeiten-Termine 2026-27.csv"
    )
    out_path = Path(
        sys.argv[2]
        if len(sys.argv) > 2
        else Path(__file__).parent / "schularbeiten-hak-steyr-2026-27.json"
    )

    school_year = "2026/27"

    rows: list[dict] = []
    with csv_path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    schularbeiten: list[dict] = []
    classes_set: dict[str, int] = {}
    subjects_set: set[str] = set()
    seq = 0

    for row in rows:
        beginn = parse_dt(row["Beginn"])
        if not beginn:
            print(f"Warnung: kein Beginn für „{row.get('Titel')}“, übersprungen.", file=sys.stderr)
            continue
        datum = beginn.strftime("%Y-%m-%d")
        beginn_uhrzeit = beginn.strftime("%H:%M")
        eh = int(row.get("Anzahl EH") or "1")
        ende_raw = (row.get("Ende") or row.get("Ende-Berechnet") or "").strip()
        ende = parse_dt(ende_raw) if ende_raw else None
        ende_uhrzeit = ende.strftime("%H:%M") if ende else ""
        dauer = dauer_aus_beginn_ende(beginn, ende, eh)
        fach = infer_fach(row)
        if not fach:
            print(f"Warnung: kein Fach für Zeile „{row.get('Titel')}“, übersprungen.", file=sys.stderr)
            continue
        subjects_set.add(fach)
        klassen = parse_json_array(row.get("Klasse") or "[]")
        if not klassen:
            print(f"Warnung: keine Klassen für „{row.get('Titel')}“, übersprungen.", file=sys.stderr)
            continue
        einheiten = parse_json_array(row.get("Einheit") or "[]")
        eh_note = ", ".join(einheiten) if einheiten else f"{eh} EH"
        titel = (row.get("Titel") or row.get("JG+Titel") or "").strip()
        sem = semester_for(beginn)
        jg = row.get("Jahrgang") or ""

        for klasse in klassen:
            klasse = str(klasse).strip()
            if not klasse:
                continue
            classes_set[klasse] = jg_year(jg) or classes_set.get(klasse, 0)
            seq += 1
            sid = f"sa-hak-steyr-{seq:04d}-{klasse.lower()}"
            schularbeiten.append(
                {
                    "Titel": titel,
                    "Thema": "",
                    "SchularbeitId": sid,
                    "FachCode": fach,
                    "KlasseCode": klasse,
                    "LehrerCode": "",
                    "LehrerEmail": "",
                    "Datum": datum,
                    "BeginnUhrzeit": beginn_uhrzeit,
                    "DauerMinuten": dauer,
                    "EndeUhrzeit": ende_uhrzeit,
                    "Semester": sem,
                    "Schuljahr": school_year,
                    "Status": "fixiert",
                    "Notiz": (
                        f"hak-steyr-2026-27 · Import · {eh_note} · "
                        f"{beginn_uhrzeit}"
                        + (f"–{ende_uhrzeit}" if ende_uhrzeit else "")
                    ),
                    "AblehnungsGrund": "",
                    "BeantragtVon": "",
                    "FixiertVon": "",
                    "FixiertAm": "",
                }
            )

    classes = [
        {"code": c, "name": c, "year": classes_set.get(c) or int(c[0]) if c and c[0].isdigit() else 0}
        for c in sorted(classes_set.keys())
    ]
    subjects = [
        {"code": c, "name": FACH_NAMES.get(c, c)}
        for c in sorted(subjects_set)
    ]
    fach_meta = []
    for i, c in enumerate(sorted(subjects_set)):
        fach_meta.append(
            {
                "Title": FACH_NAMES.get(c, c),
                "FachCode": c,
                "Farbe": FACH_COLORS[i % len(FACH_COLORS)],
                "HatSchularbeiten": True,
                "ProSemester": 2,
                "StandardDauer": 50,
                "Schuljahr": school_year,
            }
        )

    pack = {
        "schoolYear": school_year,
        "seedTag": "hak-steyr-2026-27",
        "siteDefault": "",
        "stammdaten": {
            "schoolName": "HAK Steyr",
            "domain": "",
            "subjects": subjects,
            "classes": classes,
            "teachers": [],
        },
        "regelwerk": {
            "Title": "HAK Steyr Schularbeiten 2026/27 (Import)",
            "RegelwerkId": "rw-hak-steyr-2627",
            "Schuljahr": school_year,
            "MaxProTag": 1,
            "MaxProWoche": 2,
            "AnkuendigungsfristTage": 7,
            "SperreVorNotenkonferenzTage": 7,
            "Aktiv": True,
        },
        "fachMeta": fach_meta,
        "terminfenster": [
            {
                "Title": "Weihnachtsferien 2026/27",
                "TerminfensterId": "tf-weihnacht-26",
                "Typ": "gesperrt",
                "Startdatum": "2026-12-24",
                "Enddatum": "2027-01-06",
                "Beschreibung": "hak-steyr-2026-27",
            },
            {
                "Title": "Semesterferien 2027",
                "TerminfensterId": "tf-semester-27",
                "Typ": "gesperrt",
                "Startdatum": "2027-02-01",
                "Enddatum": "2027-02-07",
                "Beschreibung": "hak-steyr-2026-27",
            },
        ],
        "schularbeiten": schularbeiten,
        "counts": {
            "schularbeiten": len(schularbeiten),
            "terminfenster": 2,
            "fachMeta": len(fach_meta),
            "teachers": 0,
            "classes": len(classes),
            "subjects": len(subjects),
        },
    }

    out_path.write_text(json.dumps(pack, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Geschrieben: {out_path}")
    print(f"  {len(rows)} CSV-Zeilen -> {len(schularbeiten)} Schularbeiten-Eintraege")
    print(f"  {len(classes)} Klassen, {len(subjects)} Fächer")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
