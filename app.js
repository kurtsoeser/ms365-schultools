(function () {
    'use strict';

    const STORAGE_KEY = 'webuntis-teams-creator-state-v1';
    const INVALID_CHARS_REPLACE = /[\\%&*+\/=?{}|<>();:,\[\]"öäü]/g;
    const INVALID_CHARS_TEST = /[\\%&*+\/=?{}|<>();:,\[\]"öäü]/;

    let rawData = [];
    let filteredData = [];
    let teamsData = [];
    /** @type {'unset'|'webuntis'|'manual'} */
    let kursteamEntryMode = 'unset';
    let currentStep = 0;
    let teacherEmailMapping = {};
    let teamsGenerated = false;
    let modalOkHandler = null;

    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const teacherUploadArea = document.getElementById('teacherUploadArea');
    const teacherFileInput = document.getElementById('teacherFileInput');
    const appModal = document.getElementById('appModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalBody = document.getElementById('modalBody');
    const modalCancel = document.getElementById('modalCancel');
    const modalOk = document.getElementById('modalOk');

    function showToast(msg) {
        const el = document.getElementById('toast');
        el.textContent = msg;
        el.classList.add('show');
        clearTimeout(showToast._t);
        showToast._t = setTimeout(() => el.classList.remove('show'), 3500);
    }

    function csvEscapeField(value) {
        const s = String(value ?? '');
        if (/[",\r\n]/.test(s)) {
            return '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
    }

    function buildCsvRow(cols) {
        return cols.map(csvEscapeField).join(',') + '\r\n';
    }

    function sanitizeGruppeForMail(g) {
        if (!g || !String(g).trim()) return '';
        let s = String(g).replace(/[_\s]+/g, '-').replace(/-+/g, '-');
        s = s.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, '');
        return s;
    }

    function buildGruppenmailBase(yearPrefix, klasseForName, fach, gruppe) {
        const km = String(klasseForName).replace(/\s+/g, '-');
        const fm = String(fach).replace(/\s+/g, '-');
        let base = `${yearPrefix}-${km}-${fm}`;
        const gs = gruppe ? sanitizeGruppeForMail(gruppe) : '';
        if (gs) base += '-' + gs;
        return base.replace(/\s+/g, '-');
    }

    function resolveDuplicateGruppenmails(teams) {
        const seen = new Map();
        let adjusted = 0;
        teams.forEach(team => {
            const base = team.gruppenmail;
            let candidate = base;
            let n = 2;
            while (seen.has(candidate)) {
                candidate = base + '-' + n;
                n++;
            }
            if (candidate !== base) {
                team.gruppenmail = candidate;
                team.mailNicknameAdjusted = true;
                adjusted++;
            } else {
                team.mailNicknameAdjusted = false;
            }
            seen.set(candidate, true);
        });
        return adjusted;
    }

    function invalidateTeams() {
        teamsData = [];
        teamsGenerated = false;
        document.getElementById('teamsTableContainer').style.display = 'none';
        document.getElementById('validationResults').style.display = 'none';
        document.getElementById('continueBtn3').style.display = 'none';
    }

    function openModal(title, bodyHtml, onOk) {
        modalTitle.textContent = title;
        modalBody.innerHTML = bodyHtml;
        modalOkHandler = onOk;
        appModal.classList.add('open');
    }

    function closeModal() {
        appModal.classList.remove('open');
        modalOkHandler = null;
        modalBody.innerHTML = '';
    }

    modalCancel.addEventListener('click', closeModal);
    modalOk.addEventListener('click', () => {
        if (typeof modalOkHandler === 'function') {
            modalOkHandler();
        }
    });
    appModal.addEventListener('click', (e) => {
        if (e.target === appModal) closeModal();
    });

    function confirmModal(title, message, onConfirm) {
        openModal(title, '<p>' + escapeHtml(message) + '</p>', () => {
            closeModal();
            onConfirm();
        });
    }

    function escapeHtml(text) {
        const d = document.createElement('div');
        d.textContent = text;
        return d.innerHTML;
    }

    function attrEscape(text) {
        return String(text ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;');
    }

    function saveStateToStorage() {
        try {
            const state = {
                rawData,
                filteredData,
                teamsData,
                teacherEmailMapping,
                teamsGenerated,
                currentStep,
                yearPrefix: document.getElementById('yearPrefix').value,
                schoolDomain:
                    typeof window.ms365GetSchoolDomainNoAt === 'function'
                        ? window.ms365GetSchoolDomainNoAt()
                        : '',
                teamSeparator: document.getElementById('teamSeparator').value,
                excludeSubjects: document.getElementById('excludeSubjects').value,
                removeDuplicates: document.getElementById('removeDuplicates').checked,
                kursteamEntryMode
            };
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
            showToast('Kursteams: Zwischenstand gespeichert.');
        } catch (e) {
            showToast('Speichern fehlgeschlagen: ' + e.message);
        }
    }

    function loadStateFromStorage() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) {
                showToast('Kein gespeicherter Stand gefunden.');
                return;
            }
            const state = JSON.parse(raw);
            rawData = state.rawData || [];
            filteredData = state.filteredData || [];
            teamsData = state.teamsData || [];
            teacherEmailMapping = state.teacherEmailMapping || {};
            teamsGenerated = !!state.teamsGenerated;
            kursteamEntryMode =
                state.kursteamEntryMode === 'manual' || state.kursteamEntryMode === 'webuntis'
                    ? state.kursteamEntryMode
                    : 'unset';

            document.getElementById('yearPrefix').value = state.yearPrefix || 'WS24';
            if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                const sd = state.schoolDomain;
                const legacy = state.emailDomain;
                if (sd !== undefined && sd !== null && String(sd).trim() !== '') {
                    window.ms365SetSchoolDomainNoAt(sd);
                } else if (legacy !== undefined && legacy !== null && String(legacy).trim() !== '') {
                    window.ms365SetSchoolDomainNoAt(
                        String(legacy)
                            .trim()
                            .replace(/^@+/, '')
                    );
                }
            }
            document.getElementById('teamSeparator').value = state.teamSeparator !== undefined ? state.teamSeparator : ' | ';
            document.getElementById('excludeSubjects').value = state.excludeSubjects !== undefined ? state.excludeSubjects : 'ORD,DIR,KV';
            document.getElementById('removeDuplicates').checked = state.removeDuplicates !== false;

            if (rawData.length) {
                document.getElementById('totalRecords').textContent = rawData.length;
                document.getElementById('uniqueSubjects').textContent = new Set(rawData.map(r => r.fach).filter(f => f)).size;
                document.getElementById('uniqueTeachers').textContent = new Set(rawData.map(r => r.lehrer).filter(l => l)).size;
                document.getElementById('importStats').style.display = 'block';
            }
            if (filteredData.length) {
                document.getElementById('filteredRecords').textContent = filteredData.length;
                document.getElementById('filterStats').style.display = 'block';
                displayFilteredData();
            }
            if (Object.keys(teacherEmailMapping).length) {
                document.getElementById('teacherCount').textContent = Object.keys(teacherEmailMapping).length;
                document.getElementById('teacherMappingInfo').style.display = 'block';
            }
            if (teamsData.length && teamsGenerated) {
                displayTeamsData();
            }

            const hasRows = (state.rawData && state.rawData.length > 0);
            const step = state.currentStep !== undefined ? state.currentStep : (hasRows ? 1 : 0);
            goToStep(step);
            showToast('Kursteams: Stand geladen.');
        } catch (e) {
            showToast('Laden fehlgeschlagen: ' + e.message);
        }
    }

    function clearStorage() {
        confirmModal('Lokalen Speicher löschen', 'Den gespeicherten Zwischenstand für Kursteams in diesem Browser wirklich löschen?', () => {
            try {
                localStorage.removeItem(STORAGE_KEY);
                showToast('Kursteams: Lokaler Speicher wurde geleert.');
            } catch (e) {
                showToast('Fehler: ' + e.message);
            }
        });
    }

    function getActivePanelMode() {
        const pw = document.getElementById('panelWebuntis');
        const pj = document.getElementById('panelJahrgang');
        const pa = document.getElementById('panelArge');
        const hidden = el => !el || window.getComputedStyle(el).display === 'none';
        if (!hidden(pw)) return 'webuntis';
        if (!hidden(pj)) return 'jahrgang';
        if (!hidden(pa)) return 'arge';
        return 'webuntis';
    }

    document.getElementById('btnSaveState').addEventListener('click', () => {
        const mode = getActivePanelMode();
        if (mode === 'jahrgang' && typeof window.ms365SaveJahrgang === 'function') {
            window.ms365SaveJahrgang();
            return;
        }
        if (mode === 'arge' && typeof window.ms365SaveArge === 'function') {
            window.ms365SaveArge();
            return;
        }
        saveStateToStorage();
    });
    document.getElementById('btnLoadState').addEventListener('click', () => {
        const mode = getActivePanelMode();
        if (mode === 'jahrgang' && typeof window.ms365LoadJahrgang === 'function') {
            window.ms365LoadJahrgang();
            return;
        }
        if (mode === 'arge' && typeof window.ms365LoadArge === 'function') {
            window.ms365LoadArge();
            return;
        }
        loadStateFromStorage();
    });
    document.getElementById('btnClearStorage').addEventListener('click', () => {
        const mode = getActivePanelMode();
        if (mode === 'jahrgang' && typeof window.ms365ClearJahrgang === 'function') {
            window.ms365ClearJahrgang();
            return;
        }
        if (mode === 'arge' && typeof window.ms365ClearArge === 'function') {
            window.ms365ClearArge();
            return;
        }
        clearStorage();
    });

    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0]);
    });

    teacherUploadArea.addEventListener('click', () => teacherFileInput.click());
    teacherUploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        teacherUploadArea.classList.add('dragover');
    });
    teacherUploadArea.addEventListener('dragleave', () => teacherUploadArea.classList.remove('dragover'));
    teacherUploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        teacherUploadArea.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) handleTeacherFile(e.dataTransfer.files[0]);
    });
    teacherFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleTeacherFile(e.target.files[0]);
    });

    function handleFile(file) {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const name = (file.name || '').toLowerCase();
                let jsonData;
                if (name.endsWith('.csv')) {
                    const buf = new Uint8Array(e.target.result);
                    let text = new TextDecoder('utf-8').decode(buf);
                    if (text.charCodeAt(0) === 0xfeff) {
                        text = text.slice(1);
                    }
                    let workbook = XLSX.read(text, { type: 'string', FS: ';' });
                    let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    jsonData = XLSX.utils.sheet_to_json(firstSheet);
                    if (!jsonData.length || Object.keys(jsonData[0] || {}).length < 2) {
                        const wb2 = XLSX.read(text, { type: 'string', FS: ',' });
                        const sh2 = wb2.Sheets[wb2.SheetNames[0]];
                        const j2 = XLSX.utils.sheet_to_json(sh2);
                        if (j2.length) {
                            jsonData = j2;
                        }
                    }
                } else {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    jsonData = XLSX.utils.sheet_to_json(firstSheet);
                }
                processImportedData(jsonData);
            } catch (error) {
                showToast('Fehler beim Lesen der Datei: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleTeacherFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const name = (file.name || '').toLowerCase();
                let jsonData;
                if (name.endsWith('.csv')) {
                    const buf = new Uint8Array(e.target.result);
                    const textUtf8 = new TextDecoder('utf-8').decode(buf);
                    jsonData = parseTeacherCsvToJsonData(textUtf8);
                    processTeacherMapping(jsonData, { replace: true });
                    let n = Object.keys(teacherEmailMapping).length;
                    if (n === 0) {
                        const manual = parseTeacherListCsvLineByLine(textUtf8);
                        if (manual.length) processTeacherMapping(manual, { replace: true });
                        n = Object.keys(teacherEmailMapping).length;
                    }
                    if (n === 0) {
                        try {
                            const text1252 = new TextDecoder('windows-1252').decode(buf);
                            jsonData = parseTeacherCsvToJsonData(text1252);
                            processTeacherMapping(jsonData, { replace: true });
                            n = Object.keys(teacherEmailMapping).length;
                            if (n === 0) {
                                const manual2 = parseTeacherListCsvLineByLine(text1252);
                                if (manual2.length) processTeacherMapping(manual2, { replace: true });
                                n = Object.keys(teacherEmailMapping).length;
                            }
                        } catch (encErr) {
                            /* windows-1252 u. a. nicht verfügbar */
                        }
                    }
                    showToast(
                        n
                            ? `Lehrer-Liste: ${n} Zuordnung(en) aus Datei geladen.`
                            : 'Datei enthielt keine gültigen Zeilen (Spalten Kürzel und E-Mail).'
                    );
                } else {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    jsonData = XLSX.utils.sheet_to_json(firstSheet);
                    processTeacherMapping(jsonData, { replace: true });
                    const n = Object.keys(teacherEmailMapping).length;
                    showToast(
                        n
                            ? `Lehrer-Liste: ${n} Zuordnung(en) aus Datei geladen.`
                            : 'Datei enthielt keine gültigen Zeilen (Spalten Kürzel und E-Mail).'
                    );
                }
            } catch (error) {
                showToast('Fehler beim Lesen der Lehrer-Datei: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    /** Gleiche Spalten wie WebUntis-Lehrerliste (Kürzel, E-Mail, …). */
    function getKuerzelFromLehrerRow(row) {
        const r = normalizeImportedRowKeys(row);
        let v = (r.Kürzel || r.Kuerzel || r.kuerzel || r.Code || r.code || r.Lehrer || '')
            .toString()
            .trim();
        if (v) return v;
        for (const k of Object.keys(r)) {
            const kl = k.toLowerCase().replace(/\s+/g, '');
            if (/^(kürzel|kuerzel|code|lehrer|abbrev)$/.test(kl) || /kuerzel|kürzel/.test(kl)) {
                const x = r[k];
                if (x != null && String(x).trim()) return String(x).trim();
            }
        }
        const vals = Object.values(r)
            .map(x => (x == null ? '' : String(x).trim()))
            .filter(Boolean);
        if (vals.length >= 2) {
            if (vals[1].includes('@') && !vals[0].includes('@')) return vals[0];
            if (vals[0].includes('@') && !vals[1].includes('@')) return vals[1];
        }
        if (vals.length >= 1 && !vals[0].includes('@')) return vals[0];
        return '';
    }

    function getEmailFromLehrerRow(row) {
        const r = normalizeImportedRowKeys(row);
        let v = (r['E-Mail'] || r.Email || r.email || r.Mail || r.mail || '').toString().trim().toLowerCase();
        if (v) return v;
        for (const k of Object.keys(r)) {
            const kl = k.toLowerCase().replace(/\s+/g, '');
            if (/^(e-?mail|email|mail)$/.test(kl) || /^e-?mail$/i.test(k)) {
                const x = r[k];
                if (x != null && String(x).trim()) return String(x).trim().toLowerCase();
            }
        }
        const vals = Object.values(r)
            .map(x => (x == null ? '' : String(x).trim()))
            .filter(Boolean);
        if (vals.length >= 2) {
            if (vals[1].includes('@')) return vals[1].toLowerCase();
            if (vals[0].includes('@')) return vals[0].toLowerCase();
        }
        if (vals.length === 1 && vals[0].includes('@')) return vals[0].toLowerCase();
        return '';
    }

    /** Wenn SheetJS die Kopfzeile nicht erkennt: Zeilen mit erstem Semikolon/Komma als Trenner. */
    function parseTeacherListCsvLineByLine(text) {
        let t = String(text || '');
        if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
        const lines = t.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        if (lines.length < 2) return [];
        const out = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i];
            let sep = line.indexOf(';');
            if (sep < 0) sep = line.indexOf(',');
            if (sep < 0) continue;
            const k = line.slice(0, sep).trim();
            const em = line.slice(sep + 1).trim();
            if (!k || !em.includes('@')) continue;
            if (/^k(ue|ü)rzel$/i.test(k.replace(/\s/g, ''))) continue;
            out.push({ Kürzel: k, 'E-Mail': em });
        }
        return out;
    }

    function parseTeacherCsvToJsonData(text) {
        let s = String(text || '');
        if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
        let workbook = XLSX.read(s, { type: 'string', FS: ';' });
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(firstSheet);
        if (!jsonData.length || Object.keys(jsonData[0] || {}).length < 2) {
            const wb2 = XLSX.read(s, { type: 'string', FS: ',' });
            const sh2 = wb2.Sheets[wb2.SheetNames[0]];
            const j2 = XLSX.utils.sheet_to_json(sh2);
            if (j2.length) jsonData = j2;
        }
        return jsonData;
    }

    /**
     * Eine Zeile aus Copy-Paste: Kürzel und E-Mail (Tab, Semikolon, |, mehrere Leerzeichen oder ein Leerzeichen).
     * @returns {{ kuerzel: string, email: string } | null}
     */
    function parseTeacherEmailPasteLine(line) {
        const t = String(line || '').trim();
        if (!t || t.startsWith('#')) return null;
        let kuerzel = '';
        let email = '';
        if (t.includes('\t')) {
            const p = t.split(/\t+/).map(s => s.trim()).filter(Boolean);
            if (p.length >= 2) {
                kuerzel = p[0];
                email = p.slice(1).join(' ').trim();
            }
        }
        if (!kuerzel && t.includes(';')) {
            const semi = t.indexOf(';');
            const left = t.slice(0, semi).trim();
            const right = t.slice(semi + 1).trim();
            if (left && right) {
                kuerzel = left;
                email = right;
            }
        }
        if (!kuerzel && /\|/.test(t)) {
            const p = t.split(/\s*\|\s*/).map(s => s.trim()).filter(Boolean);
            if (p.length >= 2) {
                kuerzel = p[0];
                email = p.slice(1).join(' ').trim();
            }
        }
        if (!kuerzel && /\s{2,}/.test(t)) {
            const p = t.split(/\s{2,}/).map(s => s.trim()).filter(Boolean);
            if (p.length >= 2) {
                kuerzel = p[0];
                email = p.slice(1).join(' ').trim();
            }
        }
        if (!kuerzel) {
            const m = t.match(/^(\S+)\s+(.+)$/);
            if (m) {
                kuerzel = m[1].trim();
                email = m[2].trim();
            }
        }
        if (!kuerzel || !email) return null;
        email = email.toLowerCase();
        if (!email.includes('@')) return null;
        return { kuerzel: kuerzel.toUpperCase(), email };
    }

    function importTeacherEmailsFromPaste() {
        const textarea = document.getElementById('teacherEmailPasteInput');
        if (!textarea) return;
        const lines = textarea.value.split(/\r?\n/);
        let added = 0;
        let invalid = 0;
        lines.forEach(line => {
            const parsed = parseTeacherEmailPasteLine(line);
            if (!parsed) {
                if (String(line).trim() && !String(line).trim().startsWith('#')) invalid++;
                return;
            }
            teacherEmailMapping[parsed.kuerzel] = parsed.email;
            added++;
        });
        document.getElementById('teacherCount').textContent = Object.keys(teacherEmailMapping).length;
        document.getElementById('teacherMappingInfo').style.display = 'block';
        if (currentStep === 2.5) updateTeacherStats();
        else displayTeacherMappingTable();
        if (added > 0) {
            showToast(
                invalid > 0
                    ? `${added} Zuordnung(en) übernommen (${invalid} Zeile(n) übersprungen).`
                    : `${added} Zuordnung(en) übernommen.`
            );
        } else if (invalid > 0) {
            showToast('Keine gültigen Zeilen – pro Zeile Kürzel und E-Mail (Tab, Semikolon oder Leerzeichen).');
        } else {
            showToast('Nichts eingefügt – bitte Zeilen mit Kürzel und E-Mail eintragen.');
        }
    }

    function downloadTeacherLehrerTemplateCsv() {
        const csv =
            '\uFEFFKürzel;E-Mail\n' +
            'MU;max.mustermann@schule.de\n' +
            'BME;anna.beispiel@schule.de\n';
        downloadBlob('Lehrerliste-Vorlage.csv', csv, 'text/csv;charset=utf-8');
        showToast('CSV-Vorlage heruntergeladen.');
    }

    function downloadTeacherLehrerTemplateXlsx() {
        if (typeof XLSX === 'undefined' || !XLSX.utils || !XLSX.writeFile) {
            showToast('Excel-Bibliothek nicht geladen – Seite neu laden.');
            return;
        }
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([
            ['Kürzel', 'E-Mail'],
            ['MU', 'max.mustermann@schule.de'],
            ['BME', 'anna.beispiel@schule.de']
        ]);
        XLSX.utils.book_append_sheet(wb, ws, 'Lehrer');
        XLSX.writeFile(wb, 'Lehrerliste-Vorlage.xlsx');
        showToast('Excel-Vorlage heruntergeladen.');
    }

    function processTeacherMapping(data, options) {
        const replace = !options || options.replace !== false;
        if (replace) teacherEmailMapping = {};
        data.forEach(row => {
            const r = normalizeImportedRowKeys(row);
            const kuerzel = getKuerzelFromLehrerRow(r).toUpperCase();
            const email = getEmailFromLehrerRow(r);
            if (!kuerzel || !email) return;
            const kNorm = kuerzel.replace(/Ü/g, 'U').replace(/ü/g, 'U');
            if (/^KUERZEL$/i.test(kNorm) && (!email || /^e-?mail$/i.test(email))) return;
            teacherEmailMapping[kuerzel] = email;
        });
        document.getElementById('teacherCount').textContent = Object.keys(teacherEmailMapping).length;
        document.getElementById('teacherMappingInfo').style.display = 'block';
        if (currentStep === 2.5) updateTeacherStats();
        else displayTeacherMappingTable();
    }

    function displayTeacherMappingTable() {
        const tbody = document.getElementById('teacherMappingBody');
        tbody.replaceChildren();
        Object.entries(teacherEmailMapping).forEach(([kuerzel, email]) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            const strong = document.createElement('strong');
            strong.textContent = kuerzel;
            td1.appendChild(strong);
            const td2 = document.createElement('td');
            td2.textContent = email;
            const td3 = document.createElement('td');
            td3.textContent = '-';
            const td4 = document.createElement('td');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger';
            btn.textContent = '❌';
            btn.addEventListener('click', () => removeTeacherMapping(kuerzel));
            td4.appendChild(btn);
            tr.append(td1, td2, td3, td4);
            tbody.appendChild(tr);
        });
    }

    function displayTeacherMappingTableWithUsage(requiredTeachers) {
        const tbody = document.getElementById('teacherMappingBody');
        tbody.replaceChildren();
        Object.entries(teacherEmailMapping).forEach(([kuerzel, email]) => {
            const isUsed = requiredTeachers.includes(kuerzel);
            const tr = document.createElement('tr');
            if (!isUsed) tr.style.opacity = '0.6';
            const td1 = document.createElement('td');
            const strong = document.createElement('strong');
            strong.textContent = kuerzel;
            td1.appendChild(strong);
            const td2 = document.createElement('td');
            td2.textContent = email;
            const td3 = document.createElement('td');
            const span = document.createElement('span');
            span.style.color = isUsed ? '#28a745' : '#6c757d';
            span.textContent = isUsed ? '✓ Wird verwendet' : 'Nicht benötigt';
            td3.appendChild(span);
            const td4 = document.createElement('td');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger';
            btn.textContent = '❌';
            btn.addEventListener('click', () => removeTeacherMapping(kuerzel));
            td4.appendChild(btn);
            tr.append(td1, td2, td3, td4);
            tbody.appendChild(tr);
        });
    }

    function toggleTeacherMapping() {
        const table = document.getElementById('teacherMappingTable');
        table.style.display = table.style.display === 'none' ? 'block' : 'none';
    }

    function clearTeacherMapping() {
        confirmModal('Zuordnungen löschen', 'Alle Lehrer-Zuordnungen wirklich löschen?', () => {
            teacherEmailMapping = {};
            document.getElementById('teacherMappingInfo').style.display = 'none';
            document.getElementById('teacherMappingTable').style.display = 'none';
            if (currentStep === 2.5) updateTeacherStats();
        });
    }

    function removeTeacherMapping(kuerzel) {
        delete teacherEmailMapping[kuerzel];
        document.getElementById('teacherCount').textContent = Object.keys(teacherEmailMapping).length;
        if (currentStep === 2.5) updateTeacherStats();
        else displayTeacherMappingTable();
        if (Object.keys(teacherEmailMapping).length === 0) {
            document.getElementById('teacherMappingInfo').style.display = 'none';
            document.getElementById('teacherMappingTable').style.display = 'none';
        }
    }

    function addTeacherMapping() {
        openModal('Lehrer-Zuordnung hinzufügen',
            '<label for="modalKuerzel">Lehrer-Kürzel</label><input type="text" id="modalKuerzel" autocomplete="off">' +
            '<label for="modalEmail">E-Mail-Adresse</label><input type="email" id="modalEmail" autocomplete="off">',
            () => {
                const k = document.getElementById('modalKuerzel').value.trim();
                const em = document.getElementById('modalEmail').value.trim().toLowerCase();
                if (!k || !em) {
                    showToast('Bitte Kürzel und E-Mail ausfüllen.');
                    return;
                }
                teacherEmailMapping[k.toUpperCase()] = em;
                document.getElementById('teacherCount').textContent = Object.keys(teacherEmailMapping).length;
                document.getElementById('teacherMappingInfo').style.display = 'block';
                closeModal();
                if (currentStep === 2.5) updateTeacherStats();
                else displayTeacherMappingTable();
            });
    }

    function normalizeImportedRowKeys(row) {
        const out = {};
        Object.keys(row).forEach(k => {
            const nk = k.replace(/^\uFEFF/, '').trim();
            out[nk] = row[k];
        });
        return out;
    }

    function splitKlassenCell(raw) {
        const s = String(raw || '').trim();
        if (!s) return [];
        return s.split(/[,;]+/).map(c => c.trim()).filter(Boolean);
    }

    function applyWebuntisRows(rows) {
        kursteamEntryMode = 'webuntis';
        rawData = rows;
        filteredData = [...rawData];
        invalidateTeams();
        document.getElementById('totalRecords').textContent = rawData.length;
        document.getElementById('uniqueSubjects').textContent = new Set(rawData.map(r => r.fach).filter(f => f)).size;
        document.getElementById('uniqueTeachers').textContent = new Set(rawData.map(r => r.lehrer).filter(l => l)).size;
        document.getElementById('importStats').style.display = 'block';
    }

    /**
     * Eine Zeile aus Copy-Paste: Lehrer, Fach, Klasse (Tab, mehrere Leerzeichen oder einfache Leerzeichen).
     * @returns {{ lehrer: string, fach: string, klasse: string } | null}
     */
    function parseWebuntisPasteLine(line) {
        const t = String(line || '').trim();
        if (!t || t.startsWith('#')) return null;
        let parts;
        if (t.includes('|')) {
            parts = t.split(/\s*\|\s*/).map(s => s.trim()).filter(Boolean);
        } else if (t.includes('\t')) {
            parts = t.split(/\t+/).map(s => s.trim()).filter(Boolean);
        } else if (/\s{2,}/.test(t)) {
            parts = t.split(/\s{2,}/).map(s => s.trim()).filter(Boolean);
        } else {
            parts = t.split(/\s+/).filter(Boolean);
        }
        if (parts.length < 3) return null;
        if (parts.length === 3) {
            return { lehrer: parts[0], fach: parts[1], klasse: parts[2] };
        }
        return {
            lehrer: parts[0],
            fach: parts[1],
            klasse: parts.slice(2).join(' ').trim()
        };
    }

    function importWebuntisFromPaste() {
        const ta = document.getElementById('webuntisPasteInput');
        const text = ta ? ta.value : '';
        const lines = String(text).split(/\r?\n/);
        const seen = new Set();
        const rows = [];
        let id = 0;
        let skipped = 0;
        let dup = 0;
        lines.forEach(line => {
            const p = parseWebuntisPasteLine(line);
            if (!p) {
                if (String(line).trim()) skipped++;
                return;
            }
            const lehrer = p.lehrer.trim();
            const fach = p.fach.trim();
            const klasse = p.klasse.trim();
            if (!lehrer || !fach || !klasse) {
                skipped++;
                return;
            }
            const key = `${lehrer.toUpperCase()}|${fach.toUpperCase()}|${klasse.toUpperCase()}`;
            if (seen.has(key)) {
                dup++;
                return;
            }
            seen.add(key);
            rows.push({
                id: id++,
                klasse,
                fach,
                lehrer,
                gruppe: '',
                original: { paste: true, line }
            });
        });
        if (!rows.length) {
            showToast('Keine gültigen Zeilen (je Zeile: Lehrer, Fach, Klasse – durch Tab oder Leerzeichen getrennt).');
            return;
        }
        applyWebuntisRows(rows);
        showToast(
            rows.length +
                ' eindeutige Zeile(n)' +
                (dup ? ', ' + dup + ' Duplikat(e) entfernt' : '') +
                (skipped ? ', ' + skipped + ' Zeile(n) übersprungen' : '') +
                '.'
        );
    }

    function processImportedData(data) {
        const rows = [];
        let id = 0;
        data.forEach(origRaw => {
            const row = normalizeImportedRowKeys(origRaw);
            const lehrer = (row.Lehrer || row.lehrer || row.Teacher || row.LehrerIn || '')
                .toString()
                .trim();
            const fach = (row.Fach || row.fach || row.Subject || row.Unterrichtsfach || '')
                .toString()
                .trim();
            const klasseRaw = (
                row['Klasse(n)'] ||
                row.Klasse ||
                row.klasse ||
                row.Class ||
                ''
            )
                .toString()
                .trim();
            const gruppe = (
                row['Schülergruppe'] ||
                row.Schülergruppe ||
                row.Gruppe ||
                row.gruppe ||
                row.Group ||
                ''
            )
                .toString()
                .trim();
            if (!lehrer || !fach) return;

            const klassenParts = splitKlassenCell(klasseRaw);
            const targets = klassenParts.length ? klassenParts : [''];

            targets.forEach(klasse => {
                rows.push({
                    id: id++,
                    klasse,
                    fach,
                    lehrer,
                    gruppe,
                    original: row
                });
            });
        });
        applyWebuntisRows(rows);
    }

    function applyFilters() {
        const excludeSubjects = document.getElementById('excludeSubjects').value
            .split(',')
            .map(s => s.trim().toUpperCase())
            .filter(s => s.length > 0);
        const removeDuplicates = document.getElementById('removeDuplicates').checked;

        let filtered = rawData.filter(row => {
            if (!row.fach || !row.lehrer) return false;
            const fach = row.fach.toUpperCase().trim();
            if (excludeSubjects.includes(fach)) return false;
            if (!row.klasse || row.klasse.trim() === '') return false;
            return true;
        });

        const originalCount = filtered.length;
        const removedByFilter = rawData.length - originalCount;

        if (removeDuplicates) {
            const seen = new Set();
            filtered = filtered.filter(row => {
                const key = `${row.klasse}-${row.fach}-${row.lehrer}-${row.gruppe}`;
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        }

        filteredData = filtered;
        invalidateTeams();
        document.getElementById('filteredRecords').textContent = filtered.length;
        document.getElementById('removedDuplicates').textContent = removedByFilter + (originalCount - filtered.length);
        document.getElementById('filterStats').style.display = 'block';
        displayFilteredData();
    }

    function displayFilteredData() {
        const tbody = document.getElementById('dataTableBody');
        tbody.replaceChildren();
        filteredData.forEach((row, index) => {
            const tr = document.createElement('tr');
            const td0 = document.createElement('td');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small btn-danger';
            btn.textContent = '❌';
            btn.addEventListener('click', () => removeRow(index));
            td0.appendChild(btn);
            const td1 = document.createElement('td');
            td1.textContent = row.klasse;
            const td2 = document.createElement('td');
            td2.textContent = row.fach;
            const td3 = document.createElement('td');
            td3.textContent = row.lehrer;
            const td4 = document.createElement('td');
            td4.textContent = row.gruppe || '-';
            tr.append(td0, td1, td2, td3, td4);
            tbody.appendChild(tr);
        });
        const hasRows = filteredData.length > 0;
        document.getElementById('dataTableContainer').style.display = hasRows ? 'block' : 'none';
        document.getElementById('continueBtn2').style.display = hasRows ? 'inline-block' : 'none';
    }

    function removeRow(index) {
        const row = filteredData[index];
        filteredData.splice(index, 1);
        if (row && row.id !== undefined && row.id !== null) {
            const ri = rawData.findIndex(r => r.id === row.id);
            if (ri >= 0) rawData.splice(ri, 1);
        }
        invalidateTeams();
        displayFilteredData();
        document.getElementById('filteredRecords').textContent = filteredData.length;
    }

    function startKursteamFromWebuntis() {
        kursteamEntryMode = 'webuntis';
        goToStep(1);
    }

    function startKursteamManual() {
        kursteamEntryMode = 'manual';
        rawData = [];
        filteredData = [];
        document.getElementById('totalRecords').textContent = '0';
        document.getElementById('uniqueSubjects').textContent = '0';
        document.getElementById('uniqueTeachers').textContent = '0';
        document.getElementById('importStats').style.display = 'none';
        const fi = document.getElementById('fileInput');
        if (fi) fi.value = '';
        invalidateTeams();
        goToStep(2);
        document.getElementById('filterStats').style.display = 'none';
        document.getElementById('dataTableContainer').style.display = 'none';
        document.getElementById('continueBtn2').style.display = 'none';
        const tbody = document.getElementById('dataTableBody');
        if (tbody) tbody.replaceChildren();
    }

    function addManualDataRow() {
        openModal('Unterrichtszeile hinzufügen',
            '<label for="manualKlasse">Klasse</label><input type="text" id="manualKlasse" autocomplete="off" placeholder="z. B. 5A">' +
            '<label for="manualFach">Fach</label><input type="text" id="manualFach" autocomplete="off" placeholder="z. B. D">' +
            '<label for="manualLehrer">Lehrkraft (Kürzel)</label><input type="text" id="manualLehrer" autocomplete="off" placeholder="z. B. MEI">' +
            '<label for="manualGruppe">Schülergruppe (optional)</label><input type="text" id="manualGruppe" autocomplete="off" placeholder="leer oder z. B. G1">',
            () => {
                const klasse = document.getElementById('manualKlasse').value.trim();
                const fach = document.getElementById('manualFach').value.trim();
                const lehrer = document.getElementById('manualLehrer').value.trim();
                const gruppe = document.getElementById('manualGruppe').value.trim();
                if (!klasse || !fach || !lehrer) {
                    showToast('Bitte Klasse, Fach und Lehrkraft ausfüllen.');
                    return;
                }
                const id = Date.now() + Math.random();
                const row = {
                    id,
                    klasse,
                    fach,
                    lehrer,
                    gruppe: gruppe || '',
                    original: {}
                };
                rawData.push(row);
                filteredData.push(row);
                kursteamEntryMode = kursteamEntryMode === 'unset' ? 'manual' : kursteamEntryMode;
                invalidateTeams();
                closeModal();
                document.getElementById('filteredRecords').textContent = filteredData.length;
                document.getElementById('filterStats').style.display = 'block';
                displayFilteredData();
            });
    }

    function resetFilters() {
        filteredData = [...rawData];
        document.getElementById('excludeSubjects').value = 'ORD,DIR,KV';
        document.getElementById('removeDuplicates').checked = true;
        applyFilters();
    }

    function generateTeamNames() {
        const yearPrefix = document.getElementById('yearPrefix').value;
        const emailDomain =
            typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
                ? window.ms365GetTeacherEmailDomainSuffix()
                : '@';
        const separator = document.getElementById('teamSeparator').value;

        teamsData = filteredData.map(row => {
            let klasseForName = row.klasse;
            if (row.klasse.includes(',')) {
                klasseForName = combineClassNames(row.klasse);
            }

            const teamName = `${yearPrefix}${separator}${klasseForName}${separator}${row.fach}`;
            const gruppenmailRaw = buildGruppenmailBase(yearPrefix, klasseForName, row.fach, row.gruppe).replace(
                /\s+/g,
                '-'
            );

            const originalGruppenmail = gruppenmailRaw;
            let gruppenmail = gruppenmailRaw.replace(INVALID_CHARS_REPLACE, '');

            let besitzer = '';
            const lehrerCode = row.lehrer.toUpperCase().trim();
            if (teacherEmailMapping[lehrerCode]) {
                besitzer = teacherEmailMapping[lehrerCode];
            } else {
                besitzer = row.lehrer.toLowerCase().trim().replace(/\s+/g, '.');
                besitzer = besitzer.replace(INVALID_CHARS_REPLACE, '');
                if (!besitzer.includes('@')) besitzer += emailDomain;
            }

            const hasInvalidChars = INVALID_CHARS_TEST.test(originalGruppenmail);
            const isValid = !hasInvalidChars && teamName && gruppenmail && besitzer && gruppenmail.length > 0;
            const mappingUsed = !!teacherEmailMapping[lehrerCode];

            return {
                teamName,
                gruppenmail,
                besitzer,
                isValid,
                error: hasInvalidChars ? 'Ungültige Zeichen in Gruppenmail' : (!isValid ? 'Unvollständige Daten' : null),
                originalClass: row.klasse,
                gruppe: row.gruppe,
                mappingUsed,
                lehrerCode: lehrerCode,
                mailNicknameAdjusted: false
            };
        });

        const dupCount = resolveDuplicateGruppenmails(teamsData);
        document.getElementById('duplicateMailAdjustments').textContent = dupCount;
        teamsGenerated = true;
        displayTeamsData();
    }

    function combineClassNames(classString) {
        const classes = classString.split(',').map(c => c.trim());
        if (classes.length === 0) return classString;
        const firstClass = classes[0];
        const jahrgang = firstClass.match(/^\d+/);
        if (!jahrgang) return classString;
        const buchstaben = classes.map(c => {
            const match = c.match(/\d+([A-Z]+)/i);
            return match ? match[1].toUpperCase() : '';
        }).filter(b => b.length > 0);
        const uniqueBuchstaben = [...new Set(buchstaben.join('').split(''))].join('');
        return jahrgang[0] + uniqueBuchstaben;
    }

    function displayTeamsData() {
        const tbody = document.getElementById('teamsTableBody');
        tbody.replaceChildren();

        const validCount = teamsData.filter(t => t.isValid).length;
        const invalidCount = teamsData.length - validCount;
        const mappedCount = teamsData.filter(t => t.mappingUsed).length;
        const dupAdj = teamsData.filter(t => t.mailNicknameAdjusted).length;
        document.getElementById('duplicateMailAdjustments').textContent = dupAdj;

        teamsData.forEach((team, index) => {
            const tr = document.createElement('tr');
            if (!team.isValid) tr.classList.add('error-row');

            const td1 = document.createElement('td');
            td1.appendChild(document.createTextNode(team.teamName));
            if (team.originalClass && team.originalClass.includes(',')) {
                td1.appendChild(document.createElement('br'));
                const small = document.createElement('small');
                small.style.color = '#6c757d';
                small.textContent = '(Original: ' + team.originalClass + ')';
                td1.appendChild(small);
            }

            const td2 = document.createElement('td');
            td2.appendChild(document.createTextNode(team.gruppenmail));
            if (team.mailNicknameAdjusted) {
                td2.appendChild(document.createElement('br'));
                const small = document.createElement('small');
                small.style.color = '#ff9800';
                small.textContent = '(Mail-Nickname wegen Duplikat angepasst)';
                td2.appendChild(small);
            }
            if (team.gruppe) {
                td2.appendChild(document.createElement('br'));
                const small = document.createElement('small');
                small.style.color = '#6c757d';
                small.textContent = 'Gruppe: ' + team.gruppe;
                td2.appendChild(small);
            }

            const td3 = document.createElement('td');
            td3.appendChild(document.createTextNode(team.besitzer));
            td3.appendChild(document.createElement('br'));
            const smallM = document.createElement('small');
            smallM.style.color = team.mappingUsed ? '#28a745' : '#ffc107';
            smallM.textContent = team.mappingUsed ? '✓ Mapping' : '⚠ Generiert (' + team.lehrerCode + ')';
            td3.appendChild(smallM);

            const td4 = document.createElement('td');
            td4.textContent = team.isValid ? '✅' : '❌ ' + (team.error || 'Fehler');

            const td5 = document.createElement('td');
            const b1 = document.createElement('button');
            b1.type = 'button';
            b1.className = 'btn btn-small';
            b1.textContent = '✏️';
            b1.addEventListener('click', () => editTeam(index));
            const b2 = document.createElement('button');
            b2.type = 'button';
            b2.className = 'btn btn-small btn-danger';
            b2.textContent = '🗑️';
            b2.addEventListener('click', () => deleteTeam(index));
            td5.append(b1, b2);

            tr.append(td1, td2, td3, td4, td5);
            tbody.appendChild(tr);
        });

        document.getElementById('validTeams').textContent = validCount;
        document.getElementById('invalidTeams').textContent = invalidCount;

        const existingWarning = document.getElementById('mappingWarning');
        if (existingWarning) existingWarning.remove();
        if (mappedCount < teamsData.length) {
            const unmappedCount = teamsData.length - mappedCount;
            const warning = document.createElement('div');
            warning.id = 'mappingWarning';
            warning.className = 'alert alert-warning';
            const strong = document.createElement('strong');
            strong.textContent = '⚠️ Achtung: ';
            warning.appendChild(strong);
            warning.appendChild(document.createTextNode(
                unmappedCount + ' Lehrer haben keine E-Mail-Zuordnung. Die E-Mail-Adressen wurden automatisch generiert.'));
            document.getElementById('validationResults').appendChild(warning);
        }

        document.getElementById('validationResults').style.display = 'block';
        document.getElementById('teamsTableContainer').style.display = 'block';
        document.getElementById('continueBtn3').style.display = 'inline-block';
    }

    function editTeam(index) {
        const team = teamsData[index];
        openModal('Team bearbeiten',
            '<label for="editName">Team-Name</label><input type="text" id="editName" value="' + attrEscape(team.teamName) + '">' +
            '<label for="editMail">Gruppenmail</label><input type="text" id="editMail" value="' + attrEscape(team.gruppenmail) + '">' +
            '<label for="editOwner">Besitzer</label><input type="email" id="editOwner" value="' + attrEscape(team.besitzer) + '">',
            () => {
                const newName = document.getElementById('editName').value.trim();
                const newMail = document.getElementById('editMail').value.trim();
                const newOwner = document.getElementById('editOwner').value.trim();
                if (!newName || !newMail || !newOwner) {
                    showToast('Bitte alle Felder ausfüllen.');
                    return;
                }
                teamsData[index] = {
                    ...team,
                    teamName: newName,
                    gruppenmail: newMail,
                    besitzer: newOwner,
                    isValid: true,
                    error: null
                };
                closeModal();
                displayTeamsData();
            });
    }

    function deleteTeam(index) {
        confirmModal('Team löschen', 'Dieses Team wirklich aus der Liste entfernen?', () => {
            teamsData.splice(index, 1);
            if (teamsData.length === 0) teamsGenerated = false;
            displayTeamsData();
        });
    }

    function addManualKursteamTeam() {
        openModal(
            'Team manuell hinzufügen',
            '<label for="addKtName">Team-Name</label><input type="text" id="addKtName" autocomplete="off" placeholder="z. B. WS26 | 1A | D">' +
                '<label for="addKtMail">Gruppenmail (Nickname)</label><input type="text" id="addKtMail" autocomplete="off" placeholder="z. B. WS26-1A-D">' +
                '<label for="addKtOwner">Besitzer (E-Mail)</label><input type="email" id="addKtOwner" autocomplete="off">',
            () => {
                const teamName = document.getElementById('addKtName').value.trim();
                const gruppenmailRaw = document.getElementById('addKtMail').value.trim();
                const besitzer = document.getElementById('addKtOwner').value.trim().toLowerCase();
                if (!teamName || !gruppenmailRaw || !besitzer) {
                    showToast('Bitte alle Felder ausfüllen.');
                    return;
                }
                const originalGruppenmail = gruppenmailRaw;
                const gruppenmail = gruppenmailRaw.replace(INVALID_CHARS_REPLACE, '');
                const hasInvalidChars = INVALID_CHARS_TEST.test(originalGruppenmail);
                const isValid = !hasInvalidChars && gruppenmail.length > 0;
                teamsData.push({
                    teamName,
                    gruppenmail,
                    besitzer,
                    isValid,
                    error: hasInvalidChars ? 'Ungültige Zeichen in Gruppenmail' : !isValid ? 'Unvollständige Daten' : null,
                    originalClass: '',
                    gruppe: '',
                    mappingUsed: true,
                    lehrerCode: '',
                    mailNicknameAdjusted: false
                });
                teamsGenerated = true;
                closeModal();
                displayTeamsData();
                showToast('Team hinzugefügt.');
            }
        );
    }

    function goToStep(step) {
        const panel = document.getElementById('panelWebuntis');
        if (!panel) return;

        if (step === 4 || step === 5) {
            const validTeams = teamsData.filter(t => t.isValid);
            if (!teamsGenerated || validTeams.length === 0) {
                showToast('Bitte zuerst unter „Teams konfigurieren“ auf „Team-Namen generieren“ klicken (mindestens ein gültiges Team).');
                step = 3;
            }
        }

        panel.querySelectorAll('.step-content').forEach(el => el.classList.remove('active'));
        panel.querySelectorAll('#panelWebuntis .steps > .step').forEach(el => {
            el.classList.remove('active');
            el.classList.remove('completed');
        });

        const contentEl = panel.querySelector('.step-content[data-step="' + step + '"]');
        const tabEl = panel.querySelector('#panelWebuntis .steps > .step[data-step="' + step + '"]');
        if (!contentEl || !tabEl) return;

        contentEl.classList.add('active');
        tabEl.classList.add('active');

        const stepOrder = [0, 1, 2, 2.5, 3, 4, 5];
        const currentIndex = stepOrder.indexOf(step);
        if (currentIndex >= 0) {
            for (let i = 0; i < currentIndex; i++) {
                const prev = panel.querySelector(
                    '#panelWebuntis .steps > .step[data-step="' + stepOrder[i] + '"]'
                );
                if (prev) prev.classList.add('completed');
            }
        }

        currentStep = step;

        const hint = document.getElementById('manualKursteamHint');
        if (hint) {
            hint.style.display = step === 2 && kursteamEntryMode === 'manual' ? 'block' : 'none';
        }

        if (step === 2.5) updateTeacherStats();
        if (step === 5) prepareCSVExport();
    }

    function updateTeacherStats() {
        const uniqueTeachers = new Set(filteredData.map(row => row.lehrer.toUpperCase().trim()));
        const teachersArray = Array.from(uniqueTeachers);
        const mappedCount = teachersArray.filter(t => teacherEmailMapping[t]).length;
        const unmappedCount = teachersArray.length - mappedCount;

        document.getElementById('uniqueTeachersNeeded').textContent = teachersArray.length;
        document.getElementById('mappedTeachers').textContent = mappedCount;
        document.getElementById('unmappedTeachers').textContent = unmappedCount;
        document.getElementById('teacherRequiredStats').style.display = 'grid';

        if (unmappedCount > 0) displayMissingTeachers(teachersArray);
        else document.getElementById('missingTeachersSection').style.display = 'none';

        if (Object.keys(teacherEmailMapping).length > 0) {
            displayTeacherMappingTableWithUsage(teachersArray);
        }
    }

    function displayMissingTeachers(allTeachers) {
        const unmappedTeachers = allTeachers.filter(t => !teacherEmailMapping[t]);
        if (unmappedTeachers.length === 0) {
            document.getElementById('missingTeachersSection').style.display = 'none';
            return;
        }
        const emailDomain =
            typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
                ? window.ms365GetTeacherEmailDomainSuffix()
                : '@';
        const tbody = document.getElementById('missingTeachersBody');
        tbody.replaceChildren();
        unmappedTeachers.forEach(kuerzel => {
            const generatedEmail = kuerzel.toLowerCase() + emailDomain;
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            const strong = document.createElement('strong');
            strong.textContent = kuerzel;
            td1.appendChild(strong);
            const td2 = document.createElement('td');
            td2.textContent = generatedEmail;
            const td3 = document.createElement('td');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-small';
            btn.textContent = '➕ Hinzufügen';
            btn.addEventListener('click', () => quickAddTeacher(kuerzel, generatedEmail));
            td3.appendChild(btn);
            tr.append(td1, td2, td3);
            tbody.appendChild(tr);
        });
        document.getElementById('missingTeachersSection').style.display = 'block';
    }

    function quickAddTeacher(kuerzel, suggestedEmail) {
        openModal('E-Mail für ' + kuerzel,
            '<label for="quickEmail">E-Mail-Adresse</label><input type="email" id="quickEmail" value="' + attrEscape(suggestedEmail) + '">',
            () => {
                const email = document.getElementById('quickEmail').value.trim().toLowerCase();
                if (!email) {
                    showToast('Bitte eine E-Mail eingeben.');
                    return;
                }
                teacherEmailMapping[kuerzel] = email;
                document.getElementById('teacherCount').textContent = Object.keys(teacherEmailMapping).length;
                closeModal();
                updateTeacherStats();
                document.getElementById('teacherMappingInfo').style.display = 'block';
            });
    }

    function prepareCSVExport() {
        const validTeams = teamsData.filter(t => t.isValid);
        document.getElementById('exportCount').textContent = validTeams.length;

        const warn = document.getElementById('step4NoTeamsWarning');
        const ready = document.getElementById('step4ReadyHint');
        const dl = document.getElementById('btnDownloadCsv');
        if (validTeams.length === 0) {
            warn.style.display = 'block';
            ready.style.display = 'none';
            dl.disabled = true;
        } else {
            warn.style.display = 'none';
            ready.style.display = 'block';
            dl.disabled = false;
        }

        let csvPreview = buildCsvRow(['TeamName', 'Gruppenmail', 'Besitzer']);
        validTeams.slice(0, 5).forEach(team => {
            csvPreview += buildCsvRow([team.teamName, team.gruppenmail, team.besitzer]);
        });
        if (validTeams.length > 5) {
            csvPreview += '... (' + (validTeams.length - 5) + ' weitere Teams)\n';
        }
        document.getElementById('csvPreview').textContent = csvPreview;
    }

    function downloadCSV() {
        const validTeams = teamsData.filter(t => t.isValid);
        if (validTeams.length === 0) {
            showToast('Keine gültigen Teams zum Exportieren.');
            return;
        }
        let csv = buildCsvRow(['TeamName', 'Gruppenmail', 'Besitzer']);
        validTeams.forEach(team => {
            csv += buildCsvRow([team.teamName, team.gruppenmail, team.besitzer]);
        });
        const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'neueteams.csv';
        link.click();
        URL.revokeObjectURL(link.href);
    }

    function copyPowerShell() {
        const script = document.getElementById('powershellScript').textContent;
        navigator.clipboard.writeText(script).then(() => {
            showToast('PowerShell-Script in die Zwischenablage kopiert.');
        });
    }

    function psEscapeSingle(s) {
        return String(s ?? '').replace(/'/g, "''");
    }

    function downloadBlob(filename, text, mime) {
        const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function buildStandaloneKursteamPs1(validTeams) {
        const stamp = new Date().toISOString();
        const rows = validTeams.map(t =>
            "    [PSCustomObject]@{ TeamName = '" +
                psEscapeSingle(t.teamName) +
                "'; Gruppenmail = '" +
                psEscapeSingle(t.gruppenmail) +
                "'; Besitzer = '" +
                psEscapeSingle(t.besitzer) +
                "' }"
        );
        const loginBlock = [
            'Write-Host ""',
            'Write-Host "=== Anmeldung bei Microsoft Teams / Microsoft 365 ===" -ForegroundColor Cyan',
            'Write-Host "Konten mit MFA: bitte Option A waehlen (Browser-Anmeldung)." -ForegroundColor Yellow',
            'Write-Host ""',
            'Write-Host " [A] Interaktive Anmeldung (empfohlen, MFA moeglich)"',
            'Write-Host " [B] Benutzername + Passwort (Get-Credential) – oft nur ohne MFA zuverlaessig"',
            'Write-Host ""',
            '$loginChoice = Read-Host "Auswahl eingeben (A oder B, Standard A)"',
            'if ($loginChoice -eq "B" -or $loginChoice -eq "b") {',
            '    $script:Ms365Cred = Get-Credential -Message "Microsoft 365 / Teams Administrator"',
            '    if ($null -eq $script:Ms365Cred) { Write-Error "Anmeldung abgebrochen."; exit 1 }',
            '    Connect-MicrosoftTeams -Credential $script:Ms365Cred',
            '} else {',
            '    Connect-MicrosoftTeams',
            '}',
            ''
        ].join('\r\n');

        const lines = [];
        lines.push('#Requires -Version 5.1');
        lines.push('# Kursteam-Anlage (Microsoft Teams, Vorlage EDU_Class)');
        lines.push('# Erzeugt in der Browser-App am ' + stamp);
        lines.push('# Daten sind unten eingebettet – keine separate CSV noetig.');
        lines.push('');
        lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
        lines.push('$ErrorActionPreference = "Continue"');
        lines.push('');
        lines.push('Write-Host ""');
        lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
        lines.push('Write-Host "  Kursteam-Erstellung (EDU_Class)"      -ForegroundColor Cyan');
        lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
        lines.push('Write-Host ""');
        lines.push('');
        lines.push('if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {');
        lines.push('    Write-Host "Installiere Modul MicrosoftTeams (einmalig)..." -ForegroundColor Yellow');
        lines.push('    Install-Module MicrosoftTeams -Scope CurrentUser -Force');
        lines.push('}');
        lines.push('Import-Module MicrosoftTeams -ErrorAction Stop');
        lines.push('');
        lines.push(loginBlock);
        lines.push('$TeamsList = @(');
        lines.push(rows.join(',\r\n'));
        lines.push(')');
        lines.push('');
        lines.push('$i = 0');
        lines.push('foreach ($Team in $TeamsList) {');
        lines.push('    $i++');
        lines.push('    try {');
        lines.push('        $null = New-Team -Template "EDU_Class" -DisplayName $Team.TeamName -MailNickname $Team.Gruppenmail -Owner $Team.Besitzer -ErrorAction Stop');
        lines.push('        Write-Host ("OK [{0}/{1}] {2}" -f $i, $TeamsList.Count, $Team.Gruppenmail) -ForegroundColor Green');
        lines.push('    }');
        lines.push('    catch {');
        lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $Team.Gruppenmail, $_.Exception.Message)');
        lines.push('    }');
        lines.push('    Start-Sleep -Seconds 2');
        lines.push('}');
        lines.push('');
        lines.push('Write-Host ""');
        lines.push('Write-Host "Fertig. Fenster schliesst nicht automatisch." -ForegroundColor Cyan');
        lines.push('Read-Host "Enter druecken zum Beenden"');
        return lines.join('\r\n');
    }

    function downloadKursteamStandalonePackage() {
        const validTeams = teamsData.filter(t => t.isValid);
        if (!validTeams.length) {
            showToast('Keine gültigen Teams – zuerst Team-Namen generieren.');
            return;
        }
        if (typeof window.ms365BuildPolyglotCmd !== 'function') {
            showToast('polyglot-cmd.js fehlt – Seite neu laden.');
            return;
        }
        const ps1 = buildStandaloneKursteamPs1(validTeams);
        const cmd = window.ms365BuildPolyglotCmd({
            title: 'Kursteam-Anlage',
            echoLine: 'Starte Kursteam-Anlage mit PowerShell ...',
            psBody: ps1
        });
        downloadBlob('Kursteam-Anlage.cmd', cmd);
        showToast('Kursteam-Anlage.cmd heruntergeladen – Doppelklick zum Start.');
    }

    function resetApp() {
        confirmModal('App zurücksetzen', 'Alle Daten in dieser Sitzung wirklich verwerfen? (Lokaler Zwischenstand bleibt, bis Sie ihn löschen.)', () => {
            location.reload();
        });
    }

    document.querySelectorAll('#panelWebuntis .steps > .step').forEach(step => {
        step.setAttribute('tabindex', '0');
        step.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                step.click();
            }
        });
        step.addEventListener('click', function () {
            const stepNum = parseFloat(this.dataset.step);
            const currentStepNum = parseFloat(currentStep);
            if (stepNum <= currentStepNum || this.classList.contains('completed')) {
                goToStep(stepNum);
            }
        });
    });

    window.goToStep = goToStep;
    window.importWebuntisFromPaste = importWebuntisFromPaste;
    window.addManualKursteamTeam = addManualKursteamTeam;
    window.startKursteamFromWebuntis = startKursteamFromWebuntis;
    window.startKursteamManual = startKursteamManual;
    window.addManualDataRow = addManualDataRow;
    window.applyFilters = applyFilters;
    window.resetFilters = resetFilters;
    window.generateTeamNames = generateTeamNames;
    window.downloadCSV = downloadCSV;
    window.copyPowerShell = copyPowerShell;
    window.resetApp = resetApp;
    window.toggleTeacherMapping = toggleTeacherMapping;
    window.clearTeacherMapping = clearTeacherMapping;
    window.addTeacherMapping = addTeacherMapping;
    window.importTeacherEmailsFromPaste = importTeacherEmailsFromPaste;
    window.downloadTeacherLehrerTemplateCsv = downloadTeacherLehrerTemplateCsv;
    window.downloadTeacherLehrerTemplateXlsx = downloadTeacherLehrerTemplateXlsx;
    window.downloadKursteamStandalonePackage = downloadKursteamStandalonePackage;

    window.ms365ShowToast = showToast;

    /**
     * Snapshot für Microsoft Graph im Browser (kursteam-graph.js).
     * @returns {{ teams: { teamName: string, gruppenmail: string, besitzer: string }[] } | null }
     */
    window.ms365GetKursteamSnapshotForGraph = function () {
        const validTeams = teamsData.filter(t => t.isValid);
        if (!validTeams.length) return null;
        return {
            teams: validTeams.map(t => ({
                teamName: t.teamName,
                gruppenmail: t.gruppenmail,
                besitzer: String(t.besitzer || '').trim()
            }))
        };
    };
})();

