# Checkliste – Schul-Tool-Upgrade / Übertragung

Referenz: `docs/tool-upgrade-vorlage.md` · `docs/rollen-vertrag.md`  
Vorbild-UI: `src/tools/schulaktivitaeten-planer/`

## 0. Kontext

- Ziel-Tool: ______________________________  
- Prefix / CSS: ______  
- Familie: [ ] approval-planer [ ] register-planer [ ] projekt-buchung [ ] ablauf-kommunikation [ ] wahl-portal [ ] betrieb [ ] it-setup  
- Portal-Personas: [ ] lehrer [ ] schueler [ ] it  
- App-Rollen (kurz): ________________________________  
- Spezialfluss: ____________________________________  
- Docs gelesen: `docs/________________.md`  
- Modus: [ ] Einzel-Upgrade [ ] Verbesserung übertragen von: __________  

---

## Welle A – Fundament (alle Familien)

- [ ] Site: `resolveSchoolToolSiteUrl` + Lizenz-Bootstrap; kein Hardcode  
- [ ] Stammdaten live + Meta; local-first bei Fehler  
- [ ] 403/SPO wipe’t Demo/Cache nicht; kein Roh-Error-Banner für Alltag  
- [ ] Ownership/Actor beim Schreiben  
- [ ] Lizenz-Gate / Feature-Catalog ok  
- [ ] Logic/State möglichst DOM-frei; sinnvolle Tests  
- [ ] IT-Setup (Listen/Flow) nicht in Lehrer-Primär-Nav  
- [ ] **SPO Live Sync:** Schema ↔ Liste; Graph Pull+Push; `spoItemId`; Demo nicht pushen; SoR = SharePoint (`docs/spo-live-sync.md`)  

---

## Welle B – UX nach Familie

### Shared Calm
- [ ] Tokens `app.css` + `planner-apps-theme`  
- [ ] Action-Banner und/oder Quiet Empty wo Alltag  
- [ ] Dashed Panels nur IT/System  
- [ ] Fach-Labels; Demo-Rollen nur `?demoRole=1`  

### approval-planer
- [ ] Nav: Lehrer flach; Staff sektioniert  
- [ ] KV als Capability wenn Konzept  
- [ ] Drawer + Guided Create  
- [ ] Filter Mehrfach + Jump-Reset  
- [ ] Shared Kalender wenn Termine zentral  
- [ ] Ops-Nacharbeit nur wenn Flag existiert  

### register-planer (z. B. Schularbeiten)
- [ ] Lehrer: Schreiben klar  
- [ ] Schüler: Read-only, keine Create-CTAs  
- [ ] Admin: Überblick/Filter  
- [ ] Kein Fake-Genehmigungsfluss  

### projekt-buchung (z. B. Projektwochen)
- [ ] Planungs-UI calm  
- [ ] Bookings/Export als klarer Schritt (nicht versteckt)  
- [ ] Fehler beim Export lokal verständlich  

### ablauf-kommunikation
- [ ] Absender-Inbox calm  
- [ ] Empfänger-Formular schlank (eigene Seite ok)  
- [ ] Keine erzwungene Planer-Sidebar  

### wahl-portal
- [ ] Phasen (offen/zu) sichtbar  
- [ ] Schüler-Primärpfad einfach  
- [ ] Admin-Auswertung getrennt  

### betrieb
- [ ] Status-Pipeline nachvollziehbar  
- [ ] Melder vs. Bearbeiter getrennt  

### it-setup
- [ ] Schrittfolge klar; Erfolgsfeedback  
- [ ] IT-Look erlaubt  

---

## Welle C – Demo & Abnahme

- [ ] Demo local-first (+ Persona passend)  
- [ ] Smoke Portal-Personas / App-Rollen des Tools  
- [ ] Wenn Verbesserung allgemein → Shared oder Vorlage aktualisiert  
- [ ] Tests grün  

---

## Übertragung (wenn Modus = Propagate)

- [ ] Shared-first geprüft  
- [ ] Gleiche Familie erledigt  
- [ ] Nahe Familien nur wo Pattern passt  
- [ ] Domäne nirgends umgebogen  

## Abnahme-Notiz

Datum: ________  
Smoke: [ ] Lehrer [ ] Schüler [ ] Leitung/Admin [ ] IT  
Offen: ________________________________________________
