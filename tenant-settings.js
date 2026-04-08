(function () {
    'use strict';

    // Kompatibilitäts-Loader: weiterhin <script src="tenant-settings.js"> möglich.
    // Lädt core + ui aus demselben Ordner nach.

    if (typeof window.ms365TenantSettingsLoad === 'function') return;

    function getBaseUrl() {
        try {
            const cs = document.currentScript;
            if (cs && cs.src) return new URL(cs.src, document.baseURI);
        } catch {
            // ignore
        }
        return null;
    }

    function loadScript(url) {
        return new Promise((resolve, reject) => {
            const el = document.createElement('script');
            el.src = String(url);
            el.defer = true;
            el.async = false;
            el.onload = () => resolve(true);
            el.onerror = () => reject(new Error('Script konnte nicht geladen werden: ' + String(url)));
            document.head.appendChild(el);
        });
    }

    (async () => {
        const base = getBaseUrl();
        if (!base) return;
        const coreUrl = new URL('tenant-settings-core.js', base);
        const uiUrl = new URL('tenant-settings-ui.js', base);
        try {
            await loadScript(coreUrl);
            await loadScript(uiUrl);
        } catch {
            // absichtlich still: wenn Dateien fehlen, sollen andere Tools nicht hart crashen
        }
    })();
})();

(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-tenant-settings-v1';
    const CURRENT_VERSION = 1;

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normCode(v) {
        return normStr(v).toUpperCase();
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function loadRaw() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return null;
            return safeJsonParse(raw);
        } catch {
            return null;
        }
    }

    function normalizeSettings(obj) {
        const o = obj && typeof obj === 'object' ? obj : {};
        const domain =
            typeof window.ms365GetSchoolDomainNoAt === 'function'
                ? window.ms365GetSchoolDomainNoAt()
                : normStr(o.domain);

        const subjectsIn = Array.isArray(o.subjects) ? o.subjects : [];
        const teachersIn = Array.isArray(o.teachers) ? o.teachers : [];
        const studentsIn = Array.isArray(o.students) ? o.students : [];

        const subjectsSeen = new Set();
        const subjects = [];
        subjectsIn.forEach((s) => {
            const code = normCode(s?.code);
            const name = normStr(s?.name);
            if (!code) return;
            const key = code.toLowerCase();
            if (subjectsSeen.has(key)) return;
            subjectsSeen.add(key);
            subjects.push({ code, name });
        });

        const teachersSeen = new Set();
        const teachers = [];
        teachersIn.forEach((t) => {
            const code = normCode(t?.code);
            const name = normStr(t?.name);
            const email = normStr(t?.email).toLowerCase();
            if (!code) return;
            const key = code.toLowerCase();
            if (teachersSeen.has(key)) return;
            teachersSeen.add(key);
            teachers.push({ code, name, email });
        });

        const students = [];
        studentsIn.forEach((s) => {
            const klasse = normStr(s?.klasse || s?.class || s?.group || s?.Klassse || s?.Klasse);
            const name = normStr(s?.name);
            const email = normStr(s?.email).toLowerCase();
            if (!klasse && !name && !email) return;
            students.push({ klasse, name, email });
        });

        return {
            version: CURRENT_VERSION,
            domain: normStr(domain),
            subjects,
            teachers,
            students
        };
    }

    function save(settings) {
        const normalized = normalizeSettings(settings);
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized));
        } catch {
            // ignore
        }
        if (typeof window.ms365SetSchoolDomainNoAt === 'function' && normalized.domain) {
            window.ms365SetSchoolDomainNoAt(normalized.domain);
        }
        return normalized;
    }

    function load() {
        const raw = loadRaw();
        const normalized = normalizeSettings(raw || {});
        return normalized;
    }

    function getTeacherEmailMap() {
        const s = load();
        const map = {};
        s.teachers.forEach((t) => {
            if (t.code && t.email) map[t.code] = t.email;
        });
        return map;
    }

    function parseDelimitedLines(text) {
        const lines = String(text || '').split(/\r\n|\n|\r/);
        const out = [];
        lines.forEach((line) => {
            const t = normStr(line);
            if (!t || t.startsWith('#')) return;
            const parts = t
                .split(/[;\t,|]/)
                .map((x) => normStr(x))
                .filter(Boolean);
            if (!parts.length) return;
            out.push(parts);
        });
        return out;
    }

    function parseLinesToSubjects(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts.slice(1).join(' '));
            if (!code) return;
            out.push({ code, name });
        });
        return out;
    }

    function parseLinesToTeachers(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const code = normCode(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!code) return;
            out.push({ code, name, email });
        });
        return out;
    }

    function parseLinesToStudents(text) {
        const out = [];
        parseDelimitedLines(text).forEach((parts) => {
            const klasse = normStr(parts[0] || '');
            const name = normStr(parts[1] || '');
            const email = normStr(parts[2] || '').toLowerCase();
            if (!klasse && !name && !email) return;
            out.push({ klasse, name, email });
        });
        return out;
    }

    function normHeaderKey(k) {
        return String(k ?? '')
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '')
            .replace(/ä/g, 'ae')
            .replace(/ö/g, 'oe')
            .replace(/ü/g, 'ue')
            .replace(/ß/g, 'ss')
            .replace(/[^a-z0-9]/g, '');
    }

    function getField(row, candidates) {
        if (!row || typeof row !== 'object') return '';
        const map = new Map();
        Object.keys(row).forEach((k) => map.set(normHeaderKey(k), row[k]));
        for (const c of candidates) {
            const v = map.get(normHeaderKey(c));
            if (v != null && String(v).trim() !== '') return String(v).trim();
        }
        return '';
    }

    function ensureXlsxReady() {
        return typeof XLSX !== 'undefined' && XLSX.utils && typeof XLSX.read === 'function';
    }

    function sheetToJsonRows(workbook) {
        const sheetName = workbook.SheetNames && workbook.SheetNames[0];
        if (!sheetName) return [];
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return [];
        return XLSX.utils.sheet_to_json(sheet, { defval: '' });
    }

    function parseCsvTextToJsonRows(text) {
        if (!ensureXlsxReady()) return [];
        let s = String(text || '');
        if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
        let wb = XLSX.read(s, { type: 'string', FS: ';' });
        let rows = sheetToJsonRows(wb);
        if (!rows.length) {
            wb = XLSX.read(s, { type: 'string', FS: ',' });
            rows = sheetToJsonRows(wb);
        }
        return rows;
    }

    function downloadXlsxTemplate(filename, aoa, sheetName) {
        if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Daten');
        XLSX.writeFile(wb, filename);
        return true;
    }

    // UI binding (optional; nur wenn Elemente existieren)
    function bindUi() {
        const form = document.getElementById('tenantSettingsForm');
        if (!form) return;

        const taSubjects = document.getElementById('tenantSubjectsLines');
        const taTeachers = document.getElementById('tenantTeachersLines');
        const teachersTbody = document.getElementById('tenantTeachersTableBody');
        const btnAddTeacherRow = document.getElementById('tenantTeachersAddRow');
        const taStudents = document.getElementById('tenantStudentsLines');
        const studentsTbody = document.getElementById('tenantStudentsTableBody');
        const btnAddStudentRow = document.getElementById('tenantStudentsAddRow');
        const fileSubjects = document.getElementById('tenantSubjectsImportFile');
        const fileTeachers = document.getElementById('tenantTeachersImportFile');
        const fileStudents = document.getElementById('tenantStudentsImportFile');
        const btnSubjectsTpl = document.getElementById('tenantSubjectsTemplateXlsx');
        const btnTeachersTpl = document.getElementById('tenantTeachersTemplateXlsx');
        const btnStudentsTpl = document.getElementById('tenantStudentsTemplateXlsx');
        const btnSave = document.getElementById('tenantSettingsSave');
        const btnReload = document.getElementById('tenantSettingsReload');
        const btnExport = document.getElementById('tenantSettingsExport');
        const fileImport = document.getElementById('tenantSettingsImportFile');
        const btnClear = document.getElementById('tenantSettingsClear');
        const summary = document.getElementById('tenantSettingsSummary');

        function setSummary(text, kind) {
            if (!summary) return;
            summary.style.display = 'block';
            summary.textContent = text;
            summary.dataset.kind = kind || 'info';
        }

        function teachersToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code)};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getTeachersFromTextarea() {
            return parseLinesToTeachers(taTeachers ? taTeachers.value : '');
        }

        function setTeachersTextareaFromRows(rows) {
            if (!taTeachers) return;
            taTeachers.value = teachersToLines(rows);
        }

        function startCellEdit(td, initialValue, onCommit) {
            const prevText = String(initialValue ?? '');
            const input = document.createElement('input');
            input.className = 'cell-editor';
            input.type = 'text';
            input.value = prevText;
            td.replaceChildren(input);
            input.focus();
            input.select();

            const commit = () => {
                const next = normStr(input.value);
                onCommit(next);
            };
            const cancel = () => {
                onCommit(prevText, { cancelled: true });
            };
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    commit();
                } else if (e.key === 'Escape') {
                    e.preventDefault();
                    cancel();
                }
            });
            input.addEventListener('blur', () => commit());
        }

        function renderTeachersTableFromTextarea() {
            if (!teachersTbody) return;
            const rows = getTeachersFromTextarea();
            teachersTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                teachersTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getTeachersFromTextarea();
                        if (!all[idx]) return renderTeachersTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setTeachersTextareaFromRows(all);
                        renderTeachersTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getTeachersFromTextarea();
                    all.splice(idx, 1);
                    setTeachersTextareaFromRows(all);
                    renderTeachersTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdEmail, tdAction);
                teachersTbody.appendChild(tr);
            });
        }

        function studentsToLines(rows) {
            return (rows || [])
                .map((x) => `${normStr(x.klasse || '')};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getStudentsFromTextarea() {
            return parseLinesToStudents(taStudents ? taStudents.value : '');
        }

        function setStudentsTextareaFromRows(rows) {
            if (!taStudents) return;
            taStudents.value = studentsToLines(rows);
        }

        function renderStudentsTableFromTextarea() {
            if (!studentsTbody) return;
            const rows = getStudentsFromTextarea();
            studentsTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = '#6c757d';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                studentsTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdClass = document.createElement('td');
                tdClass.innerHTML = `<code>${row.klasse || ''}</code>`;
                tdClass.title = 'Doppelklick zum Bearbeiten';
                tdClass.addEventListener('dblclick', () => {
                    startCellEdit(tdClass, row.klasse, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].klasse;
                        all[idx].klasse = meta && meta.cancelled ? prev : normStr(next);
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getStudentsFromTextarea();
                        if (!all[idx]) return renderStudentsTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setStudentsTextareaFromRows(all);
                        renderStudentsTableFromTextarea();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getStudentsFromTextarea();
                    all.splice(idx, 1);
                    setStudentsTextareaFromRows(all);
                    renderStudentsTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdClass, tdName, tdEmail, tdAction);
                studentsTbody.appendChild(tr);
            });
        }

        function renderFromStorage() {
            const s = load();
            if (taSubjects) {
                taSubjects.value = s.subjects.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            }
            if (taTeachers) {
                taTeachers.value = s.teachers
                    .map((x) => `${x.code};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            if (taStudents) {
                taStudents.value = (s.students || [])
                    .map((x) => `${x.klasse || ''};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            renderTeachersTableFromTextarea();
            renderStudentsTableFromTextarea();
            setSummary(
                `Gespeichert: ${s.subjects.length} Fächer, ${s.teachers.length} Lehrkräfte, ${(s.students || []).length} Schüler.`,
                'ok'
            );
        }

        function downloadJson(filename, obj) {
            const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 250);
        }

        function importFileToRows(file, onRows) {
            if (!file) return;
            if (!ensureXlsxReady()) {
                setSummary('Import: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
                return;
            }
            const name = String(file.name || '').toLowerCase();
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    let jsonRows = [];
                    if (name.endsWith('.csv')) {
                        const buf = new Uint8Array(e.target.result);
                        const tryDecoders = ['utf-8', 'windows-1252'];
                        for (const enc of tryDecoders) {
                            try {
                                const text = new TextDecoder(enc).decode(buf);
                                jsonRows = parseCsvTextToJsonRows(text);
                                if (jsonRows.length) break;
                            } catch {
                                // ignore
                            }
                        }
                    } else {
                        const data = new Uint8Array(e.target.result);
                        const wb = XLSX.read(data, { type: 'array' });
                        jsonRows = sheetToJsonRows(wb);
                    }
                    onRows(jsonRows || []);
                } catch (err) {
                    setSummary('Import fehlgeschlagen: ' + (err?.message || String(err)), 'warn');
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function importSubjectsRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'fach', 'abk', 'abkuerzung', 'abbreviation']);
                const name = getField(r, ['name', 'fachname', 'bezeichnung', 'subject', 'subjectname']);
                const c = normCode(code);
                if (!c) return;
                out.push({ code: c, name: normStr(name) });
            });
            if (taSubjects) taSubjects.value = out.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            setSummary(`Fächer importiert: ${out.length}`, 'ok');
        }

        function importTeachersRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'lehrer', 'abbrev', 'abbreviation']);
                const name = getField(r, ['name', 'lehrername', 'anzeigename', 'displayname']);
                const email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                const c = normCode(code);
                if (!c) return;
                out.push({ code: c, name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taTeachers) taTeachers.value = teachersToLines(out);
            renderTeachersTableFromTextarea();
            setSummary(`Lehrkräfte importiert: ${out.length}`, 'ok');
        }

        function importStudentsRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const klasse = getField(r, ['klasse', 'class', 'gruppe', 'group']);
                const name = getField(r, ['name', 'schueler', 'schüler', 'anzeigename', 'displayname']);
                const email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                if (!klasse && !name && !email) return;
                out.push({ klasse: normStr(klasse), name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taStudents) taStudents.value = studentsToLines(out);
            renderStudentsTableFromTextarea();
            setSummary(`Schüler importiert: ${out.length}`, 'ok');
        }

        if (btnSave) {
            btnSave.addEventListener('click', () => {
                const subjects = parseLinesToSubjects(taSubjects ? taSubjects.value : '');
                const teachers = parseLinesToTeachers(taTeachers ? taTeachers.value : '');
                const students = parseLinesToStudents(taStudents ? taStudents.value : '');
                const domain =
                    typeof window.ms365GetSchoolDomainNoAt === 'function'
                        ? window.ms365GetSchoolDomainNoAt()
                        : '';
                const saved = save({ domain, subjects, teachers, students });
                setSummary(
                    `Gespeichert: ${saved.subjects.length} Fächer, ${saved.teachers.length} Lehrkräfte, ${(saved.students || []).length} Schüler.`,
                    'ok'
                );
                renderTeachersTableFromTextarea();
                renderStudentsTableFromTextarea();
            });
        }

        if (fileSubjects) {
            fileSubjects.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importSubjectsRows(rows));
                fileSubjects.value = '';
            });
        }
        if (fileTeachers) {
            fileTeachers.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importTeachersRows(rows));
                fileTeachers.value = '';
            });
        }
        if (fileStudents) {
            fileStudents.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importStudentsRows(rows));
                fileStudents.value = '';
            });
        }

        if (btnSubjectsTpl) {
            btnSubjectsTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Faecherliste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name'],
                        ['D', 'Deutsch'],
                        ['M', 'Mathematik'],
                        ['E', 'Englisch']
                    ],
                    'Faecher'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnTeachersTpl) {
            btnTeachersTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Lehrerliste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name', 'E-Mail'],
                        ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Lehrer'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        if (btnStudentsTpl) {
            btnStudentsTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Schuelerliste-Vorlage.xlsx',
                    [
                        ['Klasse', 'Name', 'E-Mail'],
                        ['1AK', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['1AK', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Schueler'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }

        if (btnReload) {
            btnReload.addEventListener('click', () => renderFromStorage());
        }

        if (btnExport) {
            btnExport.addEventListener('click', () => {
                const s = load();
                downloadJson('tenant-einstellungen.json', s);
            });
        }

        if (btnClear) {
            btnClear.addEventListener('click', () => {
                try {
                    localStorage.removeItem(STORAGE_KEY);
                } catch {
                    // ignore
                }
                if (taSubjects) taSubjects.value = '';
                if (taTeachers) taTeachers.value = '';
                if (taStudents) taStudents.value = '';
                renderTeachersTableFromTextarea();
                renderStudentsTableFromTextarea();
                setSummary('Tenant-Grundeinstellungen gelöscht (nur lokaler Browser-Speicher).', 'warn');
            });
        }

        if (fileImport) {
            fileImport.addEventListener('change', async (e) => {
                const f = e.target.files && e.target.files[0];
                if (!f) return;
                try {
                    const text = await f.text();
                    const obj = safeJsonParse(text);
                    if (!obj) {
                        setSummary('Import fehlgeschlagen: keine gültige JSON-Datei.', 'warn');
                        return;
                    }
                    const saved = save(obj);
                    if (taSubjects) taSubjects.value = saved.subjects.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
                    if (taTeachers) taTeachers.value = saved.teachers.map((x) => `${x.code};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    if (taStudents) taStudents.value = (saved.students || []).map((x) => `${x.klasse || ''};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    renderTeachersTableFromTextarea();
                    renderStudentsTableFromTextarea();
                    setSummary(
                        `Import OK: ${saved.subjects.length} Fächer, ${saved.teachers.length} Lehrkräfte, ${(saved.students || []).length} Schüler.`,
                        'ok'
                    );
                } catch (err) {
                    setSummary('Import fehlgeschlagen: ' + (err?.message || String(err)), 'warn');
                } finally {
                    fileImport.value = '';
                }
            });
        }

        if (taTeachers) {
            taTeachers.addEventListener('input', () => renderTeachersTableFromTextarea());
        }
        if (btnAddTeacherRow) {
            btnAddTeacherRow.addEventListener('click', () => {
                const all = getTeachersFromTextarea();
                all.push({ code: '', name: '', email: '' });
                setTeachersTextareaFromRows(all);
                renderTeachersTableFromTextarea();
            });
        }

        if (taStudents) {
            taStudents.addEventListener('input', () => renderStudentsTableFromTextarea());
        }
        if (btnAddStudentRow) {
            btnAddStudentRow.addEventListener('click', () => {
                const all = getStudentsFromTextarea();
                all.push({ klasse: '', name: '', email: '' });
                setStudentsTextareaFromRows(all);
                renderStudentsTableFromTextarea();
            });
        }

        renderFromStorage();
    }

    // Public API
    window.ms365TenantSettingsLoad = load;
    window.ms365TenantSettingsSave = save;
    window.ms365TenantSettingsGetTeacherEmailMap = getTeacherEmailMap;
    window.ms365TenantSettingsParseSubjectsLines = parseLinesToSubjects;
    window.ms365TenantSettingsParseTeachersLines = parseLinesToTeachers;
    window.ms365TenantSettingsParseStudentsLines = parseLinesToStudents;

    bindUi();
})();

