import { classesToLines as classesToLinesPure } from './tenant-settings-ui-classes.js';
import {
    enrichClassesFromLinkedGroups,
    enrichClassesKvFromGraphOwners,
    collectPriorClassFieldsByCode,
    deriveClassEntryFromGroup
} from './class-list-enrich.js';
import {
    enrichSubjectsFromCatalogSources,
    subjectCatalogShrinkReport,
    formatSubjectCatalogShrinkConfirmDe,
    subjectsFromImportPayload,
    deriveSubjectEntryFromGroup
} from './subject-list-enrich.js';
import { consumeWebuntisImportPayload } from './webuntis-stammdaten-import-handoff.js';
import { pickEntraGroupsMulti } from './entra-group-picker.js';
import { mergeWebuntisImportWithExisting, summarizeWebuntisMergeResult } from './webuntis-stammdaten-wizard-logic.js';

(function () {
    'use strict';

    function normStr(v) {
        try {
            if (typeof window !== 'undefined' && window.ms365Strings && typeof window.ms365Strings.normStr === 'function') {
                return window.ms365Strings.normStr(v);
            }
        } catch {
            /* fallback */
        }
        return String(v ?? '').trim();
    }

    function dlgAlert(msg, opts) {
        if (typeof window.ms365AppDialogAlert === 'function') {
            return window.ms365AppDialogAlert(msg, opts);
        }
        window.alert(msg);
        return Promise.resolve();
    }

    function dlgConfirm(msg, opts) {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(msg, opts);
        }
        return Promise.resolve(window.confirm(msg));
    }

    function dlgPrompt(msg, def, opts) {
        if (typeof window.ms365AppDialogPrompt === 'function') {
            return window.ms365AppDialogPrompt(msg, def, opts);
        }
        return Promise.resolve(window.prompt(msg, def));
    }

    function normCode(v) {
        try {
            if (typeof window !== 'undefined' && window.ms365Strings && typeof window.ms365Strings.normCode === 'function') {
                return window.ms365Strings.normCode(v);
            }
        } catch {
            /* fallback */
        }
        return normStr(v).toUpperCase();
    }

    function escapeHtml(s) {
        try {
            if (typeof window !== 'undefined' && window.ms365Strings && typeof window.ms365Strings.escapeHtml === 'function') {
                return window.ms365Strings.escapeHtml(s);
            }
        } catch {
            /* fallback */
        }
        return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
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

    function workbookFirstSheetAoa(wb) {
        const sheetName = wb && wb.SheetNames && wb.SheetNames[0];
        const sheet = sheetName && wb.Sheets ? wb.Sheets[sheetName] : null;
        return sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) : [];
    }

    /** Maximale Spaltenzahl in den ersten Zeilen (WebUntis-CSV ist oft tab-getrennt). */
    function aoaMaxColumnCount(aoa, sampleRows) {
        const n = typeof sampleRows === 'number' ? sampleRows : 5;
        let max = 0;
        const rows = Array.isArray(aoa) ? aoa : [];
        for (let i = 0; i < rows.length && i < n; i++) {
            const row = rows[i];
            if (Array.isArray(row)) max = Math.max(max, row.length);
        }
        return max;
    }

    function csvTextToWorkbook(text) {
        let s = String(text || '');
        if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
        const delimiters = ['\t', ';', ','];
        let best = null;
        let bestCols = 0;
        delimiters.forEach(function (FS) {
            try {
                const wb = XLSX.read(s, { type: 'string', FS: FS });
                const cols = aoaMaxColumnCount(workbookFirstSheetAoa(wb));
                if (cols > bestCols) {
                    bestCols = cols;
                    best = wb;
                }
            } catch {
                /* nächstes Trennzeichen */
            }
        });
        return best || XLSX.read(s, { type: 'string', FS: ';' });
    }

    function parseCsvTextToJsonRows(text) {
        if (!ensureXlsxReady()) return [];
        return sheetToJsonRows(csvTextToWorkbook(text));
    }

    function downloadXlsxTemplate(filename, aoa, sheetName) {
        if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Daten');
        XLSX.writeFile(wb, filename);
        return true;
    }

    /** @param {{ name: string, aoa: any[][] }[]} sheets */
    function downloadXlsxMultiSheet(filename, sheets) {
        if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
        const wb = XLSX.utils.book_new();
        (sheets || []).forEach((sh) => {
            const rawName = String(sh.name || 'Daten').replace(/[:\\/?*[\]]/g, '-');
            const sn = rawName.slice(0, 31) || 'Daten';
            const ws = XLSX.utils.aoa_to_sheet(sh.aoa || []);
            XLSX.utils.book_append_sheet(wb, ws, sn);
        });
        XLSX.writeFile(wb, filename);
        return true;
    }

    /**
     * Generischer Excel/CSV-Import → JSON-Zeilen (erstes Arbeitsblatt).
     * @param {File} file
     * @param {(rows: object[]) => void} onRows
     * @param {(msg: string) => void} [onError]
     */
    function importSpreadsheetFileToJsonRows(file, onRows, onError) {
        if (!file) return;
        if (!ensureXlsxReady()) {
            if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
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
                Promise.resolve(onRows(jsonRows || [])).catch(function (err) {
                    if (onError) onError('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)));
                });
            } catch (err) {
                if (onError) onError('Import fehlgeschlagen: ' + (err?.message || String(err)));
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function exposeSubjectCatalogGuardOnWindow() {
        try {
            window.ms365SubjectCatalogShrinkReport = subjectCatalogShrinkReport;
            window.ms365FormatSubjectCatalogShrinkConfirm = formatSubjectCatalogShrinkConfirmDe;
            window.ms365SubjectsFromImportPayload = subjectsFromImportPayload;
        } catch {
            /* ignore */
        }
    }

    async function confirmSubjectCatalogIfShrinks(nextSubjects, sourceLabel) {
        const loadFn =
            typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function'
                ? window.ms365TenantSettingsLoad
                : null;
        const current = loadFn ? loadFn().subjects || [] : [];
        const report = subjectCatalogShrinkReport(current, nextSubjects || []);
        if (!report.wouldShrink) return true;
        const msg = formatSubjectCatalogShrinkConfirmDe(report, sourceLabel);
        return dlgConfirm(msg, {
            title: 'Fächerliste würde verkleinert',
            okText: 'Trotzdem übernehmen',
            cancelText: 'Abbrechen',
            danger: true
        });
    }

    /** Lehrer-Zeilen wie in der Textarea: Kürzel;Name;E-Mail pro Zeile */
    function teacherJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const vor = getField(r, ['vorname', 'forename', 'firstname']);
            const nach = getField(r, ['familienname', 'nachname', 'lastname', 'longname']);
            let code = getField(r, ['kürzel', 'kuerzel', 'code', 'lehrer', 'lehrkraft', 'abbrev', 'abbreviation']);
            if (!code && (vor || nach)) {
                code = getField(r, ['name']);
            }
            let name = getField(r, ['lehrername', 'anzeigename', 'displayname']);
            if (!name && (vor || nach)) {
                name = [vor, nach].filter(Boolean).join(' ');
            } else if (!name && !vor && !nach) {
                name = getField(r, ['name']);
            }
            let email = getField(r, ['e-mail', 'email', 'mail', 'upn', 'address.email']);
            const c = normCode(code);
            if (!c) return;

            const nameNorm = normStr(name);
            const emailNorm = normStr(email).toLowerCase();
            const nameLooksLikeEmail = nameNorm.includes('@');
            const emailLooksLikeEmail = emailNorm.includes('@');

            if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                email = nameNorm;
                name = '';
            }

            out.push({ code: c, name: normStr(name), email: normStr(email).toLowerCase() });
        });
        return out.map((x) => [x.code, x.name || '', x.email || ''].join(';')).join('\n');
    }

    function teachersRowsToAoa(rows) {
        const aoa = [['Kürzel', 'Name', 'E-Mail']];
        (rows || []).forEach((r) => {
            aoa.push([normCode(r.code || ''), normStr(r.name || ''), normStr(r.email || '').toLowerCase()]);
        });
        return aoa;
    }

    function csvEscapeField(v) {
        const s = String(v ?? '');
        if (/[;"\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
    }

    function downloadTeachersCsv(filename, rows) {
        try {
            const aoa = teachersRowsToAoa(rows);
            const body = aoa.map((line) => line.map(csvEscapeField).join(';')).join('\r\n');
            const blob = new Blob(['\ufeff' + body], { type: 'text/csv;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename || 'Lehrerliste.csv';
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 250);
            return true;
        } catch {
            return false;
        }
    }

    function compareTeachersDe(a, b) {
        return String(a ?? '').localeCompare(String(b ?? ''), 'de', { sensitivity: 'base' });
    }

    /** @param {'code'|'name'|'email'} key @param {1|-1} dir */
    function sortTeacherRows(rows, key, dir) {
        const d = dir === -1 ? -1 : 1;
        const k = key === 'email' || key === 'name' ? key : 'code';
        return (rows || []).slice().sort((a, b) => {
            let av = a && a[k];
            let bv = b && b[k];
            if (k === 'code') {
                av = normCode(av);
                bv = normCode(bv);
            } else if (k === 'email') {
                av = normStr(av).toLowerCase();
                bv = normStr(bv).toLowerCase();
            } else {
                av = normStr(av);
                bv = normStr(bv);
            }
            return compareTeachersDe(av, bv) * d;
        });
    }

    function schoolDomainForEmail() {
        if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
            return normStr(window.ms365GetSchoolDomainNoAt()).replace(/^@+/, '');
        }
        return '';
    }

    function emailPatternFromSelect(id, fallback) {
        const el = document.getElementById(id);
        const v = el ? normStr(el.value) : '';
        return v || fallback || 'vorname.nachname';
    }

    /** @returns {'first'|'all'} */
    function emailFirstNameModeFromSelect(id, fallback) {
        const el = document.getElementById(id);
        const raw = el ? normStr(el.value).toLowerCase() : '';
        const api = window.ms365PersonEmailFromName;
        if (api && typeof api.resolveFirstNameMode === 'function') {
            return api.resolveFirstNameMode(raw || fallback || 'first');
        }
        if (raw === 'all' || raw === 'alle') return 'all';
        if (raw === 'first' || raw === 'erster') return 'first';
        const fb = String(fallback || 'first').toLowerCase();
        return fb === 'all' ? 'all' : 'first';
    }

    /**
     * Liest eine oder mehrere Spreadsheet-Dateien → [{ name, aoa }].
     * @param {File|FileList|File[]} files
     * @returns {Promise<{ name: string, aoa: any[][] }[]>}
     */
    function readFilesToAoaSheets(files) {
        const list = files && files.length != null ? Array.from(files) : files ? [files] : [];
        if (!ensureXlsxReady()) return Promise.reject(new Error('Excel-Bibliothek nicht geladen'));
        return Promise.all(
            list.map(function (file) {
                return new Promise(function (resolve, reject) {
                    const name = String(file.name || '').toLowerCase();
                    const reader = new FileReader();
                    reader.onload = function (e) {
                        try {
                            let wb;
                            if (name.endsWith('.csv') || name.endsWith('.txt')) {
                                wb = csvTextToWorkbook(String(e.target.result || ''));
                            } else {
                                wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                            }
                            const sheetName = wb.SheetNames && wb.SheetNames[0];
                            const sheet = sheetName ? wb.Sheets[sheetName] : null;
                            const aoa = sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) : [];
                            resolve({ name: file.name || sheetName || 'sheet', aoa: aoa });
                        } catch (err) {
                            reject(err);
                        }
                    };
                    reader.onerror = function () {
                        reject(new Error('Datei konnte nicht gelesen werden: ' + (file.name || '')));
                    };
                    if (name.endsWith('.csv') || name.endsWith('.txt')) reader.readAsText(file);
                    else reader.readAsArrayBuffer(file);
                });
            })
        );
    }

    /**
     * WebUntis- oder SIS-Import aus einer/mehreren Dateien.
     * @returns {Promise<object>} SIS-ähnliches Result { source, records, lines, meta, ... }
     */
    function importStudentSheetsResult(sheets, sourceHint, existingStudents, patternId, firstNameMode) {
        const wu = window.ms365WebuntisExportImport;
        const sis = window.ms365SchoolSisImport;
        const domain = schoolDomainForEmail();
        const pattern = patternId || 'vorname.nachname';
        const givenMode = firstNameMode || 'first';
        const hint = String(sourceHint || 'auto').toLowerCase();

        let classified = null;
        if (wu && typeof wu.classifySheets === 'function') {
            classified = wu.classifySheets(sheets);
        }

        if (
            classified &&
            (classified.studentAoa || classified.guardianAoa) &&
            (hint === 'webuntis' || hint === 'auto' || classified.studentAoa)
        ) {
            if (classified.studentAoa && wu.importStudentsFromWebuntis) {
                let result = wu.importStudentsFromWebuntis({
                    studentAoa: classified.studentAoa,
                    guardianAoa: classified.guardianAoa || [],
                    domain: domain,
                    pattern: pattern,
                    firstNameMode: givenMode,
                    applyEmails: !!domain
                });
                if (sis && typeof sis.preferExistingStudentEmails === 'function') {
                    result.records = sis.preferExistingStudentEmails(result.records, existingStudents || []);
                    result.lines = sis.recordsToSemicolonLines(result.records);
                }
                if (!domain) {
                    result.meta = result.meta || {};
                    result.meta.emailHint = 'Schuldomain in den Stammdaten setzen, um fehlende Schüler-Mails vorzuschlagen.';
                }
                return result;
            }
        }

        // Einzeldatei / klassischer SIS-Pfad
        const first = sheets && sheets[0];
        if (!first || !sis) {
            return { source: hint, records: [], lines: '', meta: { studentCount: 0, withParents: 0, parentMails: 0 } };
        }
        const objectRows = [];
        const headers = first.aoa[0] || [];
        for (let i = 1; i < first.aoa.length; i++) {
            const row = first.aoa[i] || [];
            const o = {};
            headers.forEach(function (h, idx) {
                const key = normStr(h);
                if (!key) return;
                if (o[key] == null || o[key] === '') o[key] = row[idx];
            });
            objectRows.push(o);
        }
        let result = sis.importStudentsAndGuardians({
            aoa: first.aoa,
            objectRows: objectRows,
            source: hint
        });
        if (domain && wu && typeof wu.applyPersonEmails === 'function') {
            const needMail = (result.records || []).some(function (r) {
                return !normStr(r.email);
            });
            if (needMail) {
                const em = wu.applyPersonEmails(result.records, {
                    domain: domain,
                    pattern: pattern,
                    firstNameMode: givenMode
                });
                result.records = em.records;
                result.lines = sis.recordsToSemicolonLines(result.records);
                result.emailMeta = { generated: em.generated, conflicts: em.conflicts };
                result.meta = Object.assign({}, result.meta, {
                    emailsGenerated: em.generated,
                    emailConflicts: em.conflicts
                });
            }
        }
        if (sis && typeof sis.preferExistingStudentEmails === 'function') {
            result.records = sis.preferExistingStudentEmails(result.records, existingStudents || []);
            result.lines = sis.recordsToSemicolonLines(result.records);
        }
        return result;
    }

    window.ms365TeacherListImport = {
        isXlsxReady: ensureXlsxReady,
        sortRows: sortTeacherRows,
        downloadTemplate() {
            return downloadXlsxTemplate(
                'Lehrerliste-Vorlage.xlsx',
                [
                    ['Kürzel', 'Name', 'E-Mail'],
                    ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                    ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                ],
                'Lehrer'
            );
        },
        downloadCsvTemplate() {
            return downloadTeachersCsv('Lehrerliste-Vorlage.csv', [
                { code: 'MU', name: 'Max Mustermann', email: 'max.mustermann@schule.de' }
            ]);
        },
        exportCsv(rows, filename) {
            return downloadTeachersCsv(filename || 'Lehrerliste.csv', rows);
        },
        exportXlsx(rows, filename) {
            return downloadXlsxTemplate(filename || 'Lehrerliste.xlsx', teachersRowsToAoa(rows), 'Lehrer');
        },
        importFile(file, onLines, onError) {
            if (!file) return;
            readFilesToAoaSheets([file])
                .then(function (sheets) {
                    const wu = window.ms365WebuntisExportImport;
                    const aoa = sheets[0] && sheets[0].aoa;
                    const kind =
                        wu && aoa && typeof wu.detectExportKindFromAoa === 'function' ? wu.detectExportKindFromAoa(aoa) : '';
                    if (kind === 'teacher' && wu.importTeachersFromWebuntis) {
                        const domain = schoolDomainForEmail();
                        const pattern = emailPatternFromSelect('tenantTeachersEmailPattern', emailPatternFromSelect('swTeachersEmailPattern'));
                        const firstNameMode = emailFirstNameModeFromSelect(
                            'tenantTeachersEmailGivenNames',
                            emailFirstNameModeFromSelect('swTeachersEmailGivenNames')
                        );
                        const result = wu.importTeachersFromWebuntis({
                            teacherAoa: aoa,
                            domain: domain,
                            pattern: pattern,
                            firstNameMode: firstNameMode,
                            applyEmails: !!domain
                        });
                        if (onLines) onLines(result.lines, result);
                        return;
                    }
                    const rows = sheets[0]
                        ? XLSX.utils.sheet_to_json(
                              XLSX.utils.aoa_to_sheet(sheets[0].aoa),
                              { defval: '' }
                          )
                        : [];
                    // TeacherSalary-ähnliche Header ohne Erkennung: Fallback JSON
                    if (onLines) onLines(teacherJsonRowsToSemicolonLines(rows));
                })
                .catch(function (err) {
                    if (onError) onError('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)));
                });
        }
    };

    /** Schüler-Zeilen: Klasse;Name;E-Mail[;Eltern1;Eltern1Mail;…] */
    function studentJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const klasse = getField(r, ['klasse', 'class', 'zug', 'gruppe', 'k']);
            let name = getField(r, ['name', 'schueler', 'schüler', 'vollname', 'displayname']);
            let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
            const nameNorm = normStr(name);
            const emailNorm = normStr(email).toLowerCase();
            const nameLooksLikeEmail = nameNorm.includes('@');
            const emailLooksLikeEmail = emailNorm.includes('@');
            if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                email = nameNorm;
                name = '';
            }
            const k = normStr(klasse);
            if (!k && !normStr(name) && !normStr(email)) return;
            const parts = [k, normStr(name), normStr(email).toLowerCase()];
            const e1 = getField(r, ['eltern1', 'erziehungsberechtigte1', 'mutter', 'parent1', 'guardian1']);
            const e1m = getField(r, ['eltern1mail', 'eltern1-mail', 'eltern1email', 'parent1email', 'guardian1email', 'muttermail']);
            const e2 = getField(r, ['eltern2', 'erziehungsberechtigte2', 'vater', 'parent2', 'guardian2']);
            const e2m = getField(r, ['eltern2mail', 'eltern2-mail', 'eltern2email', 'parent2email', 'guardian2email', 'vatermail']);
            if (normStr(e1m).includes('@') || normStr(e1).includes('@')) {
                parts.push(normStr(e1));
                parts.push(normStr(e1m || (String(e1).includes('@') ? e1 : '')).toLowerCase());
            }
            if (normStr(e2m).includes('@') || normStr(e2).includes('@')) {
                parts.push(normStr(e2));
                parts.push(normStr(e2m || (String(e2).includes('@') ? e2 : '')).toLowerCase());
            }
            out.push(parts.join(';'));
        });
        return out.join('\n');
    }

    function studentsRowsToAoa(rows) {
        const aoa = [['Klasse', 'Name', 'E-Mail', 'Eltern1', 'Eltern1Mail', 'Eltern2', 'Eltern2Mail']];
        (rows || []).forEach((r) => {
            const pairs = Array.isArray(r.parentPairs) ? r.parentPairs : [];
            const p1 = pairs[0] || {};
            const p2 = pairs[1] || {};
            aoa.push([
                normStr(r.klasse || ''),
                normStr(r.name || ''),
                normStr(r.email || '').toLowerCase(),
                normStr(p1.name || ''),
                normStr(p1.email || '').toLowerCase(),
                normStr(p2.name || ''),
                normStr(p2.email || '').toLowerCase()
            ]);
        });
        return aoa;
    }

    function downloadStudentsCsv(filename, rows) {
        try {
            const aoa = studentsRowsToAoa(rows);
            const body = aoa.map((line) => line.map(csvEscapeField).join(';')).join('\r\n');
            const blob = new Blob(['\ufeff' + body], { type: 'text/csv;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename || 'Schuelerliste.csv';
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 250);
            return true;
        } catch {
            return false;
        }
    }

    /** @param {'klasse'|'name'|'email'} key @param {1|-1} dir */
    function sortStudentRows(rows, key, dir) {
        const d = dir === -1 ? -1 : 1;
        const k = key === 'email' || key === 'name' ? key : 'klasse';
        return (rows || []).slice().sort((a, b) => {
            let av = a && a[k];
            let bv = b && b[k];
            if (k === 'email') {
                av = normStr(av).toLowerCase();
                bv = normStr(bv).toLowerCase();
            } else {
                av = normStr(av);
                bv = normStr(bv);
            }
            return compareTeachersDe(av, bv) * d;
        });
    }

    function adminJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const role = getField(r, ['rolle', 'role', 'position', 'funktion']);
            const name = getField(r, ['name', 'anzeigename', 'displayname']);
            const email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
            if (!normStr(role) && !normStr(name) && !normStr(email)) return;
            out.push({ role: normStr(role), name: normStr(name), email: normStr(email).toLowerCase() });
        });
        return out.map((x) => [x.role, x.name || '', x.email || ''].filter(Boolean).join(';')).join('\n');
    }

    function subjectsJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = getField(r, ['kürzel', 'kuerzel', 'code', 'fach', 'abbrev']);
            const name = getField(r, ['name', 'bezeichnung', 'fachname', 'displayname']);
            const c = normCode(code);
            if (!c) return;
            out.push({ code: c, name: normStr(name) });
        });
        return out.map((x) => [x.code, x.name || ''].join(';')).join('\n');
    }

    function argesJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = getField(r, ['kürzel', 'kuerzel', 'code', 'arge', 'kuerzelarge']);
            const name = getField(r, ['name', 'bezeichnung', 'displayname']);
            const subj = getField(r, ['fächer', 'faecher', 'subjects', 'fachzuordnung', 'faecherkuerzel']);
            const c = normCode(code);
            if (!c) return;
            let line = c + ';' + normStr(name);
            if (normStr(subj)) line += ';' + normStr(subj);
            out.push(line);
        });
        return out.join('\n');
    }

    function classesJsonRowsToSemicolonLines(jsonRows) {
        const out = [];
        (jsonRows || []).forEach((r) => {
            const code = getField(r, ['kürzel', 'kuerzel', 'code', 'klassekurz']);
            const year = getField(r, ['abschlussjahr', 'jahr', 'year', 'jahrgang']);
            const name = normStr(getField(r, ['anzeigename', 'klasse', 'name', 'displayname']));
            const headName = normStr(getField(r, ['kvname', 'klassenvorstand', 'vorstand', 'headname']));
            const headEmail = normStr(getField(r, ['kvmail', 'kv-email', 'e-mailkv', 'heademail', 'email'])).toLowerCase();
            const c = normCode(code);
            const y = /^\d{4}$/.test(normStr(year)) ? normStr(year) : '';
            if (!c && !name && !y && !headName && !headEmail) return;
            if (y) {
                out.push([c, y, name, headName, headEmail].join(';'));
            } else {
                out.push([c, name, headName, headEmail].join(';'));
            }
        });
        return out.join('\n');
    }

    function findWorksheetName(wb, aliases) {
        const names = wb.SheetNames || [];
        const want = (aliases || []).map((a) => normHeaderKey(a));
        for (const sn of names) {
            const nk = normHeaderKey(sn);
            if (want.indexOf(nk) >= 0) return sn;
        }
        for (const sn of names) {
            const nk = normHeaderKey(sn);
            for (const w of want) {
                if (!w) continue;
                if (nk === w || nk.indexOf(w) === 0 || w.indexOf(nk) === 0) return sn;
            }
        }
        return null;
    }

    function importFileToWorkbook(file, onWorkbook, onError) {
        if (!file) return;
        if (!ensureXlsxReady()) {
            if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
            return;
        }
        const name = String(file.name || '').toLowerCase();
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                let wb;
                if (name.endsWith('.csv') || name.endsWith('.txt')) {
                    wb = csvTextToWorkbook(String(e.target.result || ''));
                } else {
                    wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                }
                if (onWorkbook) onWorkbook(wb);
            } catch (err) {
                if (onError) onError('Import fehlgeschlagen: ' + (err?.message || String(err)));
            }
        };
        if (name.endsWith('.csv') || name.endsWith('.txt')) reader.readAsText(file, 'utf-8');
        else reader.readAsArrayBuffer(file);
    }

    function workbookToSheetList(wb, fileName) {
        const names = (wb && wb.SheetNames) || [];
        const out = [];
        names.forEach(function (sn) {
            const sheet = wb.Sheets[sn];
            if (!sheet) return;
            const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            if (!aoa || !aoa.length) return;
            out.push({
                name: (fileName ? fileName + ' — ' : '') + sn,
                aoa: aoa
            });
        });
        return out;
    }

    function readFilesToAllSheets(files) {
        const list = files && files.length != null ? Array.from(files) : files ? [files] : [];
        if (!ensureXlsxReady()) return Promise.reject(new Error('Excel-Bibliothek nicht geladen'));
        return Promise.all(
            list.map(function (file) {
                return new Promise(function (resolve, reject) {
                    const name = String(file.name || '').toLowerCase();
                    const reader = new FileReader();
                    reader.onload = function (e) {
                        try {
                            let wb;
                            if (name.endsWith('.csv') || name.endsWith('.txt')) {
                                wb = csvTextToWorkbook(String(e.target.result || ''));
                            } else {
                                wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                            }
                            resolve(workbookToSheetList(wb, file.name || ''));
                        } catch (err) {
                            reject(err);
                        }
                    };
                    reader.onerror = function () {
                        reject(new Error('Datei konnte nicht gelesen werden: ' + (file.name || '')));
                    };
                    if (name.endsWith('.csv') || name.endsWith('.txt')) reader.readAsText(file);
                    else reader.readAsArrayBuffer(file);
                });
            })
        ).then(function (nested) {
            const flat = [];
            nested.forEach(function (arr) {
                (arr || []).forEach(function (sh) {
                    flat.push(sh);
                });
            });
            return flat;
        });
    }

    window.ms365SpreadsheetCsv = {
        csvTextToWorkbook: csvTextToWorkbook
    };

    window.ms365StammdatenFileIO = {
        readFilesToAoaSheets: readFilesToAoaSheets,
        readFilesToAllSheets: readFilesToAllSheets
    };

    window.ms365StudentListImport = {
        isXlsxReady: ensureXlsxReady,
        sortRows: sortStudentRows,
        downloadTemplate() {
            if (window.ms365SchoolSisImport && typeof window.ms365SchoolSisImport.downloadXlsxTemplates === 'function') {
                if (window.ms365SchoolSisImport.downloadXlsxTemplates()) return true;
            }
            return downloadXlsxTemplate(
                'Schuelerliste-Vorlage.xlsx',
                [
                    ['Klasse', 'Name', 'E-Mail', 'Eltern1', 'Eltern1Mail', 'Eltern2', 'Eltern2Mail'],
                    ['1AK', 'Lisa Beispiel', 'lisa.beispiel@schule.de', 'Maria Beispiel', 'maria@mail.com', '', ''],
                    ['1AK', 'Max Muster', 'max.muster@schule.de', 'Eva Muster', 'eva@mail.com', 'Tom Muster', 'tom@mail.com']
                ],
                'Schueler'
            );
        },
        downloadCsvTemplate() {
            if (window.ms365SchoolSisImport && typeof window.ms365SchoolSisImport.downloadCsv === 'function') {
                return window.ms365SchoolSisImport.downloadCsv(
                    'Schueler-Eltern-Vorlage.csv',
                    window.ms365SchoolSisImport.ms365TemplateAoa()
                );
            }
            return downloadStudentsCsv('Schuelerliste-Vorlage.csv', [
                {
                    klasse: '1AK',
                    name: 'Lisa Beispiel',
                    email: 'lisa.beispiel@schule.de',
                    parentPairs: [{ name: 'Maria Beispiel', email: 'maria@mail.com' }]
                }
            ]);
        },
        exportCsv(rows, filename) {
            return downloadStudentsCsv(filename || 'Schuelerliste.csv', rows);
        },
        exportXlsx(rows, filename) {
            return downloadXlsxTemplate(filename || 'Schuelerliste.xlsx', studentsRowsToAoa(rows), 'Schueler');
        },
        importFile(file, onLines, onError, sourceHint, existingStudents) {
            const files = file && file.length != null ? file : file ? [file] : [];
            if (!files.length) return;
            if (!ensureXlsxReady()) {
                if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
                return;
            }
            const pattern = emailPatternFromSelect(
                'tenantStudentsEmailPattern',
                emailPatternFromSelect('swStudentsEmailPattern', emailPatternFromSelect('evImportEmailPattern'))
            );
            const firstNameMode = emailFirstNameModeFromSelect(
                'tenantStudentsEmailGivenNames',
                emailFirstNameModeFromSelect(
                    'swStudentsEmailGivenNames',
                    emailFirstNameModeFromSelect('evImportEmailGivenNames')
                )
            );
            readFilesToAoaSheets(files)
                .then(function (sheets) {
                    const result = importStudentSheetsResult(
                        sheets,
                        sourceHint || 'auto',
                        existingStudents || [],
                        pattern,
                        firstNameMode
                    );
                    if (onLines) onLines(result.lines, result);
                })
                .catch(function (err) {
                    if (onError) onError('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)));
                });
        },
        importFiles(files, onLines, onError, sourceHint, existingStudents) {
            return this.importFile(files, onLines, onError, sourceHint, existingStudents);
        }
    };

    window.ms365SchuldatenMasterImport = {
        isXlsxReady: ensureXlsxReady,
        downloadTemplate() {
            return downloadXlsxMultiSheet('MS365-Schuldaten-Vorlage.xlsx', [
                {
                    name: 'Verwaltung',
                    aoa: [
                        ['Rolle', 'Name', 'E-Mail'],
                        ['Direktion', 'Direktorin Beispiel', 'direktion@schule.de'],
                        ['Sekretariat', 'Sekretariat', 'sekretariat@schule.de']
                    ]
                },
                {
                    name: 'Lehrer',
                    aoa: [
                        ['Kürzel', 'Name', 'E-Mail'],
                        ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ]
                },
                {
                    name: 'Schueler',
                    aoa: [
                        ['Klasse', 'Name', 'E-Mail', 'Eltern1', 'Eltern1Mail', 'Eltern2', 'Eltern2Mail'],
                        ['1AK', 'Lisa Beispiel', 'lisa.beispiel@schule.de', 'Maria Beispiel', 'maria@mail.com', '', ''],
                        ['1AK', 'Max Muster', 'max.muster@schule.de', 'Eva Muster', 'eva@mail.com', 'Tom Muster', 'tom@mail.com']
                    ]
                },
                {
                    name: 'Faecher',
                    aoa: [
                        ['Kürzel', 'Name'],
                        ['D', 'Deutsch'],
                        ['M', 'Mathematik'],
                        ['E', 'Englisch']
                    ]
                },
                {
                    name: 'ARGE',
                    aoa: [
                        ['Kürzel', 'Name', 'Fächer'],
                        ['SPRACHEN', 'Sprachen', 'D,E'],
                        ['NAWI', 'Naturwissenschaften', 'BIO,CH,PH']
                    ]
                },
                {
                    name: 'Klassen',
                    aoa: [
                        ['Kürzel', 'Abschlussjahr', 'Anzeigename', 'KV-Name', 'KV-E-Mail'],
                        ['HMA', '2031', '1HMA', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['1AK', '2030', '1A-Klasse', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ]
                }
            ]);
        },
        importFile(file, onPayload, onError) {
            const name = String(file.name || '').toLowerCase();
            if (name.endsWith('.csv')) {
                if (onError) onError('Gesamt-Import: Bitte die XLSX-Vorlage verwenden (mehrere Arbeitsblätter).');
                return;
            }
            importFileToWorkbook(
                file,
                (wb) => {
                    try {
                        const out = {};
                        const snV = findWorksheetName(wb, ['verwaltung', 'administration']);
                        const snL = findWorksheetName(wb, ['lehrer', 'lehrerinnen', 'teachers']);
                        const snS = findWorksheetName(wb, ['schueler', 'schuler', 'schüler', 'students', 'schuelerinnen']);
                        const snF = findWorksheetName(wb, ['faecher', 'fächer', 'subjects']);
                        const snA = findWorksheetName(wb, ['arge', 'arbeitsgruppen', 'arbeitsgemeinschaften']);
                        const snK = findWorksheetName(wb, ['klassen', 'classes']);
                        function sheetLines(sn, conv) {
                            if (!sn || !wb.Sheets[sn]) return null;
                            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: '' });
                            if (!rows || !rows.length) return null;
                            const lines = conv(rows);
                            return normStr(lines).length ? lines : null;
                        }
                        const v = sheetLines(snV, adminJsonRowsToSemicolonLines);
                        if (v != null) out.verwaltungLines = v;
                        const l = sheetLines(snL, teacherJsonRowsToSemicolonLines);
                        if (l != null) out.lehrerLines = l;
                        const s = sheetLines(snS, studentJsonRowsToSemicolonLines);
                        if (s != null) out.schuelerLines = s;
                        const f = sheetLines(snF, subjectsJsonRowsToSemicolonLines);
                        if (f != null) out.faecherLines = f;
                        const a = sheetLines(snA, argesJsonRowsToSemicolonLines);
                        if (a != null) out.argeLines = a;
                        const k = sheetLines(snK, classesJsonRowsToSemicolonLines);
                        if (k != null) out.klassenLines = k;
                        if (onPayload) onPayload(out);
                    } catch (err) {
                        if (onError) onError(String(err?.message || err));
                    }
                },
                onError
            );
        }
    };

    // UI binding (optional; nur wenn Elemente existieren)
    function bindUi() {
        exposeSubjectCatalogGuardOnWindow();
        const form = document.getElementById('tenantSettingsForm');
        if (!form) return;

        if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
            return;
        }

        const parseLinesToSubjects = window.ms365TenantSettingsParseSubjectsLines;
        const parseLinesToArges = window.ms365TenantSettingsParseArgesLines;
        const parseLinesToTeachers = window.ms365TenantSettingsParseTeachersLines;
        const parseLinesToAdmin = window.ms365TenantSettingsParseAdminLines;
        const parseLinesToAdminRoles = window.ms365TenantSettingsParseAdminRolesLines;
        const parseLinesToStudents = window.ms365TenantSettingsParseStudentsLines;
        const parseLinesToStudentCouncil = window.ms365TenantSettingsParseStudentCouncilLines;
        const parseLinesToClasses = window.ms365TenantSettingsParseClassesLines;
        const parseLinesToSga = window.ms365TenantSettingsParseSgaLines;
        const load = window.ms365TenantSettingsLoad;
        const save = window.ms365TenantSettingsSave;

        const taSubjects = document.getElementById('tenantSubjectsLines');
        const subjectsTbody = document.getElementById('tenantSubjectsTableBody');
        const subjectsTable = subjectsTbody ? subjectsTbody.closest('table') : null;
        const subjectsSortState = { key: null, dir: 1 };
        const btnAddSubjectRow = document.getElementById('tenantSubjectsAddRow');
        const btnSubjectsImportM365 = document.getElementById('tenantSubjectsImportM365');
        const btnEnrichSubjects = document.getElementById('tenantBtnEnrichSubjects');
        const taArges = document.getElementById('tenantArgesLines');
        const argesTbody = document.getElementById('tenantArgesTableBody');
        const argesTable = argesTbody ? argesTbody.closest('table') : null;
        const argesSortState = { key: null, dir: 1 };
        const btnAddArgeRow = document.getElementById('tenantArgesAddRow');
        const taTeachers = document.getElementById('tenantTeachersLines');
        const teachersTbody = document.getElementById('tenantTeachersTableBody');
        const teachersTable = teachersTbody ? teachersTbody.closest('table') : null;
        const btnAddTeacherRow = document.getElementById('tenantTeachersAddRow');
        const btnVerifyTeachersGraph = document.getElementById('tenantBtnVerifyTeachersGraph');
        const btnTeachersExportCsv = document.getElementById('tenantTeachersExportCsv');
        const btnTeachersExportXlsx = document.getElementById('tenantTeachersExportXlsx');
        /** @type {{ key: 'code'|'name'|'email'|null, dir: 1|-1 }} */
        const teachersSortState = { key: null, dir: 1 };
        const taAdminBundle = document.getElementById('tenantAdminBundleLines');
        const taAdmin = document.getElementById('tenantAdminLines');
        const adminTbody = document.getElementById('tenantAdminTableBody');
        const adminUnifiedTbody = document.getElementById('tenantAdminUnifiedTableBody');
        const adminUnifiedTable = adminUnifiedTbody ? adminUnifiedTbody.closest('table') : null;
        const adminUnifiedSortState = { key: null, dir: 1 };
        const btnAddAdminRow = document.getElementById('tenantAdminAddRow');
        const taAdminRoles = document.getElementById('tenantAdminRoleLines');
        const adminRolesTbody = document.getElementById('tenantAdminRolesTableBody');
        const btnAdminRolesDefaults = document.getElementById('tenantAdminRolesDefaults');
        const btnVerifyVerwaltungGraph = document.getElementById('tenantBtnVerifyVerwaltungGraph');
        const selSgaMode = document.getElementById('tenantSgaMode');
        const taSga = document.getElementById('tenantSgaLines');
        const sgaTbody = document.getElementById('tenantSgaTableBody');
        const sgaTable = sgaTbody ? sgaTbody.closest('table') : null;
        const sgaSortState = { key: null, dir: 1 };
        const btnAddSgaRow = document.getElementById('tenantSgaAddRow');
        const btnVerifySgaGraph = document.getElementById('tenantBtnVerifySgaGraph');
        const sgaGroupMatchCell = document.getElementById('tenantSgaGroupMatchCell');
        const btnVerifySgaGroup = document.getElementById('tenantBtnVerifySgaGroup');
        const btnCreateSgaGroup = document.getElementById('tenantBtnCreateSgaGroup');
        const btnUnmatchSgaGroup = document.getElementById('tenantBtnUnmatchSgaGroup');
        const btnOpenSgaGroup = document.getElementById('tenantBtnOpenSgaGroup');
        const inpSgaGroupSearch = document.getElementById('tenantSgaGroupSearch');
        const btnSearchSgaGroup = document.getElementById('tenantBtnSearchSgaGroup');
        const sgaGroupSearchResults = document.getElementById('tenantSgaGroupSearchResults');
        const ENTRA_GROUP_DETAILS =
            'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/';
        const taStudents = document.getElementById('tenantStudentsLines');
        const studentsTbody = document.getElementById('tenantStudentsTableBody');
        const studentsTable = studentsTbody ? studentsTbody.closest('table') : null;
        const studentsClassFilter = document.getElementById('tenantStudentsClassFilter');
        const studentsTextFilter = document.getElementById('tenantStudentsTextFilter');
        const studentsFilterMeta = document.getElementById('tenantStudentsFilterMeta');
        const btnAddStudentRow = document.getElementById('tenantStudentsAddRow');
        const btnVerifyStudentsGraph = document.getElementById('tenantBtnVerifyStudentsGraph');
        const btnStudentsExportCsv = document.getElementById('tenantStudentsExportCsv');
        const btnStudentsExportXlsx = document.getElementById('tenantStudentsExportXlsx');
        /** @type {{ key: 'klasse'|'name'|'email'|null, dir: 1|-1 }} */
        const studentsSortState = { key: null, dir: 1 };
        const taStudentCouncil = document.getElementById('tenantStudentCouncilLines');
        const studentCouncilTbody = document.getElementById('tenantStudentCouncilTableBody');
        const studentCouncilTable = studentCouncilTbody ? studentCouncilTbody.closest('table') : null;
        const studentCouncilSortState = { key: null, dir: 1 };
        const btnAddStudentCouncilRow = document.getElementById('tenantStudentCouncilAddRow');
        const btnVerifyStudentCouncilGraph = document.getElementById('tenantBtnVerifyStudentCouncilGraph');
        const studentCouncilGroupMatchCell = document.getElementById('tenantStudentCouncilGroupMatchCell');
        const btnVerifyStudentCouncilGroup = document.getElementById('tenantBtnVerifyStudentCouncilGroup');
        const btnCreateStudentCouncilGroup = document.getElementById('tenantBtnCreateStudentCouncilGroup');
        const btnUnmatchStudentCouncilGroup = document.getElementById('tenantBtnUnmatchStudentCouncilGroup');
        const btnOpenStudentCouncilGroup = document.getElementById('tenantBtnOpenStudentCouncilGroup');
        const inpStudentCouncilGroupSearch = document.getElementById('tenantStudentCouncilGroupSearch');
        const btnSearchStudentCouncilGroup = document.getElementById('tenantBtnSearchStudentCouncilGroup');
        const studentCouncilGroupSearchResults = document.getElementById('tenantStudentCouncilGroupSearchResults');
        const taClasses = document.getElementById('tenantClassesLines');
        const classesTbody = document.getElementById('tenantClassesTableBody');
        const classesTable = classesTbody ? classesTbody.closest('table') : null;
        const classesSortState = { key: null, dir: 1 };
        const btnAddClassRow = document.getElementById('tenantClassesAddRow');
        const btnVerifyClassesGraph = document.getElementById('tenantBtnVerifyClassesGraph');
        const btnClassesImportM365 = document.getElementById('tenantClassesImportM365');
        const btnEnrichClassesFromGroups = document.getElementById('tenantBtnEnrichClassesFromGroups');
        const fileSubjects = document.getElementById('tenantSubjectsImportFile');
        const fileArges = document.getElementById('tenantArgesImportFile');
        const btnArgesImportM365 = document.getElementById('tenantArgesImportM365');
        const fileTeachers = document.getElementById('tenantTeachersImportFile');
        const fileStudents = document.getElementById('tenantStudentsImportFile');
        const fileClasses = document.getElementById('tenantClassesImportFile');
        const btnSubjectsTpl = document.getElementById('tenantSubjectsTemplateXlsx');
        const btnArgesTpl = document.getElementById('tenantArgesTemplateXlsx');
        const btnTeachersTpl = document.getElementById('tenantTeachersTemplateXlsx');
        const btnStudentsTpl = document.getElementById('tenantStudentsTemplateXlsx');
        const btnClassesTpl = document.getElementById('tenantClassesTemplateXlsx');
        const btnSave = document.getElementById('tenantSettingsSave');
        const btnReload = document.getElementById('tenantSettingsReload');
        const btnExport = document.getElementById('tenantSettingsExport');
        const btnExportHeader = document.getElementById('tenantSettingsExportHeader');
        const fileImport = document.getElementById('tenantSettingsImportFile');
        const btnClear = document.getElementById('tenantSettingsClear');
        const summary = document.getElementById('tenantSettingsSummary');
        const _inpDefaultGradYear = null;
        const schoolNameInput = document.getElementById('schoolName');
        const domainInput = document.getElementById('schoolEmailDomain');
        const schoolYearSelect = document.getElementById('schoolYearSelect');
        const schoolYearAddBtn = document.getElementById('schoolYearAddBtn');

        /** Schuljahr Sep–Aug (kanonisch, siehe shared/utils/school-year.js). */
        function currentSchoolYearLabel(now) {
            try {
                if (typeof window !== 'undefined' && window.ms365SchoolYear && typeof window.ms365SchoolYear.currentSchoolYearLabel === 'function') {
                    return window.ms365SchoolYear.currentSchoolYearLabel(now);
                }
            } catch {
                /* fallback */
            }
            const d = now instanceof Date && !isNaN(now.getTime()) ? now : new Date();
            const y = d.getMonth() < 8 ? d.getFullYear() - 1 : d.getFullYear();
            return String(y) + '/' + String(y + 1).slice(2);
        }

        function getDisplayedSchoolYearLabel() {
            try {
                if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                    const c = window.ms365AppDataV2.getContainer();
                    const cur = c && c.years ? String(c.years.current || '').trim() : '';
                    if (cur) return cur;
                }
            } catch {
                // ignore
            }
            if (schoolYearSelect) return String(schoolYearSelect.value || '').trim();
            return '';
        }

        let autoSaveTimer = null;
        let __syncGuard = 0;

        function dispatchTenantSettingsChanged(saved, reason) {
            try {
                if (__syncGuard) return;
                window.dispatchEvent(
                    new CustomEvent('ms365-tenant-settings-changed', {
                        detail: { settings: saved, reason: String(reason || '') }
                    })
                );
            } catch {
                // ignore
            }
        }

        function autoSaveNow() {
            const subjects = typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
            const arges = typeof parseLinesToArges === 'function' ? parseLinesToArges(taArges ? taArges.value : '') : [];
            const teachers = typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
            const admin = getAdminFromTextarea();
            const adminRoles = getAdminRolesFromTextarea();
            const sga = typeof parseLinesToSga === 'function' ? parseLinesToSga(taSga ? taSga.value : '') : [];
            const sgaMode = selSgaMode ? normStr(selSgaMode.value || 'group').toLowerCase() : 'group';
            const students = typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
            const studentCouncil =
                typeof parseLinesToStudentCouncil === 'function'
                    ? parseLinesToStudentCouncil(taStudentCouncil ? taStudentCouncil.value : '')
                    : [];
            const classes = typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
            const schoolName = schoolNameInput ? normStr(schoolNameInput.value || '') : '';
            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
            const administration = getAdministrationEntries();
            const saved = save({ schoolName, domain, subjects, arges, teachers, administration, admin, adminRoles, sgaMode, sga, students, studentCouncil, classes });
            dispatchTenantSettingsChanged(saved, 'autosave');
        }

        function argesToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const list = Array.isArray(x.subjects) ? x.subjects : [];
                    return `${normCode(x.code)};${normStr(x.name || '')};${list.map((s) => normCode(s)).filter(Boolean).join(',')}`.trim();
                })
                .filter(Boolean)
                .join('\n');
        }

        function getArgesFromTextarea() {
            return typeof parseLinesToArges === 'function' ? parseLinesToArges(taArges ? taArges.value : '') : [];
        }

        function setArgesTextareaFromRows(rows) {
            if (!taArges) return;
            taArges.value = argesToLines(rows);
        }

        function getArgeCatalogLink(code) {
            const api = window.ms365AppDataV2;
            if (api && typeof api.getCatalogLink === 'function') return api.getCatalogLink('arge', code);
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const links = (setup && setup.catalogLinks) || [];
            const c = normCode(code);
            return links.find(function (x) {
                return x && x.kind === 'arge' && normCode(x.code) === c;
            });
        }

        function upsertArgeCatalogLink(code, group) {
            const api = window.ms365AppDataV2;
            if (!api || typeof api.upsertCatalogLink !== 'function') return;
            api.upsertCatalogLink({
                kind: 'arge',
                code: normCode(code),
                graphGroupId: group && group.id ? String(group.id) : '',
                displayName: group && group.displayName ? String(group.displayName) : '',
                mailNickname: group && group.mailNickname ? String(group.mailNickname) : '',
                mode: 'matched'
            });
        }

        function getSubjectCatalogLink(code) {
            const api = window.ms365AppDataV2;
            if (api && typeof api.getCatalogLink === 'function') return api.getCatalogLink('subject', code);
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const links = (setup && setup.catalogLinks) || [];
            const c = normCode(code);
            return links.find(function (x) {
                return x && x.kind === 'subject' && normCode(x.code) === c;
            });
        }

        function upsertSubjectCatalogLink(code, group) {
            const api = window.ms365AppDataV2;
            if (!api || typeof api.upsertCatalogLink !== 'function') return;
            api.upsertCatalogLink({
                kind: 'subject',
                code: normCode(code),
                graphGroupId: group && group.id ? String(group.id) : '',
                displayName: group && group.displayName ? String(group.displayName) : '',
                mailNickname: group && group.mailNickname ? String(group.mailNickname) : '',
                mode: 'matched'
            });
        }

        function subjectGroupMailPrefixForDerive() {
            const api = window.ms365AppDataV2;
            const su = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            return normStr(su && su.subjectGroupMailPrefix) || 'fach';
        }

        function deriveArgeEntryFromGroup(g) {
            const nick = normStr(g && g.mailNickname).toLowerCase();
            const dn = normStr(g && g.displayName);
            const prefixes = ['arge-', 'arge.', 'fachgruppe-', 'fachgruppe.', 'fg-', 'fg.', 'ag-', 'ag.'];
            let rest = nick;
            for (let i = 0; i < prefixes.length; i++) {
                if (rest.startsWith(prefixes[i])) {
                    rest = rest.slice(prefixes[i].length);
                    break;
                }
            }
            let code = rest ? normCode(rest.replace(/[^a-z0-9]/gi, '')) : '';
            if (!code && dn) {
                const m = dn.match(/^(?:ARGE|Fachgruppe|FG|AG)\s*[-–:]?\s*(.+)$/i);
                if (m) {
                    const token = normStr(m[1]).split(/\s+/)[0] || '';
                    code = normCode(token);
                }
            }
            if (!code && nick) code = normCode(nick.replace(/[^a-z0-9]/gi, '').slice(0, 20));
            const name = dn || code || normStr(g && g.mailNickname);
            return { code, name };
        }

        function buildArgeM365Cell(row) {
            const td = document.createElement('td');
            td.style.fontSize = '0.88em';
            td.style.lineHeight = '1.35';
            const code = normCode(row && row.code);
            if (!code) {
                td.style.color = 'var(--muted)';
                td.textContent = '–';
                td.title = 'Kürzel nötig für die Verknüpfung';
                return td;
            }
            const link = getArgeCatalogLink(code);
            if (link && normStr(link.graphGroupId)) {
                const nick = normStr(link.mailNickname) || normStr(link.displayName);
                td.innerHTML =
                    '<span style="color:#0d8050;font-weight:700;">✓</span> ' +
                    escapeHtml(nick || 'verknüpft');
                td.title =
                    (link.displayName ? String(link.displayName) : '') +
                    (link.mailNickname ? '\nAlias: ' + String(link.mailNickname) : '') +
                    '\nGroup-ID: ' +
                    String(link.graphGroupId);
            } else {
                td.style.color = 'var(--muted)';
                td.textContent = '–';
                td.title = 'Noch nicht mit einer Microsoft-365-Gruppe verknüpft';
            }
            return td;
        }

        async function importArgesFromM365Groups() {
            try {
                const searchUnified =
                    graphApi() && typeof graphApi().searchUnifiedGroups === 'function'
                        ? async function (q, token) {
                              const tok = token || (await graphApi().getGraphToken());
                              return graphApi().searchUnifiedGroups(tok, q);
                          }
                        : undefined;
                const picked = await pickEntraGroupsMulti({
                    title: 'ARGE-Gruppen aus Microsoft 365',
                    hint:
                        'Es werden Vorschläge mit den Suchbegriffen arge, fg und ag geladen. Ausgewählte Gruppen werden als ARGE-Stammdaten übernommen und in catalogLinks verknüpft (wie in der Einrichtung).',
                    initialQueries: ['arge', 'fg', 'ag'],
                    searchFn: searchUnified
                });
                if (!picked.length) return;
                const all = getArgesFromTextarea();
                const byCode = new Map();
                all.forEach(function (r) {
                    const c = normCode(r.code);
                    if (c) byCode.set(c, r);
                });
                let added = 0;
                let linked = 0;
                let skipped = 0;
                picked.forEach(function (g) {
                    const derived = deriveArgeEntryFromGroup(g);
                    const code = normCode(derived.code);
                    if (!code) {
                        skipped++;
                        return;
                    }
                    if (!byCode.has(code)) {
                        byCode.set(code, { code, name: derived.name || code, subjects: [] });
                        added++;
                    } else {
                        const ex = byCode.get(code);
                        if (!normStr(ex.name) && derived.name) ex.name = derived.name;
                    }
                    upsertArgeCatalogLink(code, g);
                    linked++;
                });
                const next = Array.from(byCode.values()).sort(function (a, b) {
                    return normCode(a.code).localeCompare(normCode(b.code));
                });
                setArgesTextareaFromRows(next);
                renderArgesTableFromTextarea();
                scheduleAutoSave();
                const parts = [];
                if (added) parts.push(added + ' neue ARGE');
                if (linked) parts.push(linked + ' verknüpft');
                if (skipped) parts.push(skipped + ' übersprungen (kein Kürzel)');
                setSummary(
                    'Microsoft 365: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'),
                    added || linked ? 'ok' : 'warn'
                );
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                setSummary('Import aus Microsoft 365: ' + msg, 'warn');
            }
        }

        function buildSubjectM365Cell(row) {
            const td = document.createElement('td');
            td.style.fontSize = '0.88em';
            td.style.lineHeight = '1.35';
            const code = normCode(row && row.code);
            if (!code) {
                td.style.color = 'var(--muted)';
                td.textContent = '–';
                td.title = 'Kürzel nötig für die Verknüpfung';
                return td;
            }
            const link = getSubjectCatalogLink(code);
            if (link && normStr(link.graphGroupId)) {
                const nick = normStr(link.mailNickname) || normStr(link.displayName);
                td.innerHTML =
                    '<span style="color:#0d8050;font-weight:700;">✓</span> ' +
                    escapeHtml(nick || 'verknüpft');
                td.title =
                    (link.displayName ? String(link.displayName) : '') +
                    (link.mailNickname ? '\nAlias: ' + String(link.mailNickname) : '') +
                    '\nGroup-ID: ' +
                    String(link.graphGroupId);
            } else {
                td.style.color = 'var(--muted)';
                td.textContent = '–';
                td.title = 'Noch nicht mit einer Microsoft-365-Fachgruppe verknüpft';
            }
            return td;
        }

        async function importSubjectsFromM365Groups() {
            try {
                const searchUnified =
                    graphApi() && typeof graphApi().searchUnifiedGroups === 'function'
                        ? async function (q, token) {
                              const tok = token || (await graphApi().getGraphToken());
                              return graphApi().searchUnifiedGroups(tok, q);
                          }
                        : undefined;
                const prefix = subjectGroupMailPrefixForDerive();
                const picked = await pickEntraGroupsMulti({
                    title: 'Fachgruppen aus Microsoft 365',
                    hint:
                        'Vorschläge mit fach, fachgruppe und Ihrem Präfix („' +
                        prefix +
                        '“). Ausgewählte Gruppen werden als Fachgruppen-Stammdaten übernommen und in catalogLinks verknüpft (wie ARGE & Fachgruppen).',
                    initialQueries: [prefix, 'fach', 'fachgruppe'],
                    searchFn: searchUnified
                });
                if (!picked.length) return;
                const all = getSubjectsFromTextarea();
                const byCode = new Map();
                all.forEach(function (r) {
                    const c = normCode(r.code);
                    if (c) byCode.set(c, r);
                });
                let added = 0;
                let linked = 0;
                let skipped = 0;
                const deriveOpts = { mailPrefix: prefix };
                picked.forEach(function (g) {
                    const derived = deriveSubjectEntryFromGroup(g, deriveOpts);
                    const code = normCode(derived.code);
                    if (!code) {
                        skipped++;
                        return;
                    }
                    if (!byCode.has(code)) {
                        byCode.set(code, { code, name: derived.name || code });
                        added++;
                    } else {
                        const ex = byCode.get(code);
                        if (!normStr(ex.name) && derived.name) ex.name = derived.name;
                    }
                    upsertSubjectCatalogLink(code, g);
                    linked++;
                });
                const next = Array.from(byCode.values()).sort(function (a, b) {
                    return normCode(a.code).localeCompare(normCode(b.code));
                });
                setSubjectsTextareaFromRows(next);
                renderSubjectsTableFromTextarea();
                scheduleAutoSave();
                const parts = [];
                if (added) parts.push(added + ' neue Fachgruppe(n)');
                if (linked) parts.push(linked + ' verknüpft');
                if (skipped) parts.push(skipped + ' übersprungen (kein Kürzel)');
                setSummary(
                    'Microsoft 365: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'),
                    added || linked ? 'ok' : 'warn'
                );
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                setSummary('Import aus Microsoft 365: ' + msg, 'warn');
            }
        }

        function renderArgesTableFromTextarea() {
            if (!argesTbody) return;
            updateTableSortIndicators(argesTable, argesSortState);
            const rows = getArgesFromTextarea();
            argesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 5;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                argesTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.textContent = row.code || '';
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getArgesFromTextarea();
                        if (!all[idx]) return renderArgesTableFromTextarea();
                        const prev = all[idx].code;
                        const newCode = meta && meta.cancelled ? prev : normCode(next);
                        if (prev && newCode && normCode(prev) !== normCode(newCode)) {
                            const api = window.ms365AppDataV2;
                            if (api && typeof api.renameCatalogLink === 'function') {
                                try {
                                    api.renameCatalogLink('arge', prev, newCode);
                                } catch {
                                    /* Verknüpfung bleibt am alten Kürzel */
                                }
                            }
                        }
                        all[idx].code = newCode;
                        setArgesTextareaFromRows(all);
                        renderArgesTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getArgesFromTextarea();
                        if (!all[idx]) return renderArgesTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setArgesTextareaFromRows(all);
                        renderArgesTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdSubjects = document.createElement('td');
                tdSubjects.textContent = (Array.isArray(row.subjects) ? row.subjects : []).join(', ');
                tdSubjects.title = 'Doppelklick zum Bearbeiten';
                tdSubjects.addEventListener('dblclick', () => {
                    startCellEdit(tdSubjects, (Array.isArray(row.subjects) ? row.subjects : []).join(','), (next, meta) => {
                        const all = getArgesFromTextarea();
                        if (!all[idx]) return renderArgesTableFromTextarea();
                        const prev = all[idx].subjects;
                        all[idx].subjects =
                            meta && meta.cancelled
                                ? prev
                                : String(next || '')
                                      .split(/[,\s|]+/)
                                      .map((x) => normCode(x))
                                      .filter(Boolean);
                        setArgesTextareaFromRows(all);
                        renderArgesTableFromTextarea();
                        scheduleAutoSave();
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
                    const all = getArgesFromTextarea();
                    all.splice(idx, 1);
                    setArgesTextareaFromRows(all);
                    renderArgesTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                const tdM365 = buildArgeM365Cell(row);

                tr.append(tdCode, tdName, tdSubjects, tdM365, tdAction);
                argesTbody.appendChild(tr);
            });
        }

        function scheduleAutoSave() {
            if (autoSaveTimer) clearTimeout(autoSaveTimer);
            autoSaveTimer = setTimeout(() => {
                autoSaveTimer = null;
                try {
                    autoSaveNow();
                } catch {
                    // ignore (z.B. während Import/Reset)
                }
            }, 450);
        }

        function setSummary(text, kind) {
            if (!summary) return;
            summary.style.display = 'block';
            summary.textContent = text;
            summary.dataset.kind = kind || 'info';
        }

        // ----------------------------------------------------------------
        // Status-Übersicht: „Was fehlt noch?"
        // ----------------------------------------------------------------
        function renderStatusOverview() {
            const grid = document.getElementById('tenantStatusGrid');
            const footer = document.getElementById('tenantStatusFooter') || document.getElementById('tenantStatusLastCheck');
            if (!grid) return;

            const api = window.ms365AppDataV2;
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : {};
            const dirMatch = (setup && setup.directoryMatchByEmail) ? setup.directoryMatchByEmail : {};
            const classMatch = (setup && setup.classGroupMatchByKey) ? setup.classGroupMatchByKey : {};

            // Alle Personen-E-Mails aus allen Listen einsammeln
            const teacherRows = getTeachersFromTextarea ? getTeachersFromTextarea() : [];
            const adminRows = typeof getAdministrationGroups === 'function'
                ? (function () {
                    const gs = getAdministrationGroups();
                    const out = [];
                    (gs || []).forEach(function (g) {
                        (Array.isArray(g.people) ? g.people : []).forEach(function (p) {
                            if (p && p.email) out.push({ name: p.name || '', email: p.email });
                        });
                    });
                    return out;
                }())
                : [];
            const studentRows = getStudentsFromTextarea ? getStudentsFromTextarea() : [];
            const sgaRows = typeof getSgaFromTextarea === 'function' ? getSgaFromTextarea() : [];
            const svRows = typeof getStudentCouncilFromTextarea === 'function' ? getStudentCouncilFromTextarea() : [];

            function countEmailMatch(rows) {
                let total = 0; let matched = 0; let notFound = 0; let unchecked = 0;
                const seen = new Set();
                rows.forEach(function (r) {
                    const em = normStr(r && r.email || '').toLowerCase();
                    if (!em || em.indexOf('@') === -1) return;
                    if (seen.has(em)) return;
                    seen.add(em);
                    total++;
                    const m = dirMatch[em];
                    if (m && m.graphUserId) matched++;
                    else if (m && m.notFound) notFound++;
                    else unchecked++;
                });
                return { total, matched, notFound, unchecked };
            }

            const tStat = countEmailMatch(teacherRows);
            const aStat = countEmailMatch(adminRows);
            const sStat = countEmailMatch(studentRows);
            const sgaStat = countEmailMatch(sgaRows);
            const svStat = countEmailMatch(svRows);

            // Klassen-Gruppen
            const classRows = typeof getClassesFromTextarea === 'function' ? getClassesFromTextarea() : [];
            let clTotal = 0; let clMatched = 0; let clNotFound = 0; let clUnchecked = 0;
            const clSeen = new Set();
            classRows.forEach(function (row) {
                const key = (function () {
                    const code = normStr(row && row.code || '').toUpperCase();
                    const name = normStr(row && row.name || '').toUpperCase();
                    return code || name || '';
                }());
                if (!key || clSeen.has(key)) return;
                clSeen.add(key);
                clTotal++;
                const m = classMatch[key];
                if (m && m.groupId) clMatched++;
                else if (m && m.notFound) clNotFound++;
                else clUnchecked++;
            });

            // Fächer/ARGEs (ohne Gruppen-Match-Check: einfach zählen, wie viele vorhanden)
            const subjectRows = typeof getSubjectsFromTextarea === 'function' ? getSubjectsFromTextarea() : [];
            const argeRows = typeof getArgesFromTextarea === 'function' ? getArgesFromTextarea() : [];

            // Chips aufbauen
            const chips = [];

            function chip(icon, label, kind, title) {
                chips.push({ icon, label, kind, title: title || '' });
            }

            function personChip(label, stat, _tab) {
                const { total, matched, notFound, unchecked } = stat;
                if (total === 0) {
                    chip('bi-dash', label + ': keine Einträge', 'muted', 'Keine Einträge in der Liste.');
                    return;
                }
                if (matched === total) {
                    chip('bi-check-circle-fill', label + ': alle ' + total + ' gematcht', 'ok', 'Alle ' + total + ' E-Mail-Adressen in Microsoft Entra gefunden.');
                } else if (unchecked > 0) {
                    chip('bi-question-circle', label + ': ' + unchecked + '/' + total + ' ungeprüft', 'warn',
                        unchecked + ' noch nicht geprüft, ' + notFound + ' nicht gefunden, ' + matched + ' gefunden.');
                } else if (notFound > 0) {
                    chip('bi-x-circle', label + ': ' + notFound + '/' + total + ' nicht gefunden', 'error',
                        notFound + ' in Microsoft Entra nicht gefunden, ' + matched + ' gefunden.');
                } else {
                    chip('bi-check-circle-fill', label + ': alle ' + total + ' gematcht', 'ok', 'Alle gefunden.');
                }
            }

            personChip('Lehrer', tStat);
            personChip('Verwaltung', aStat);
            personChip('Schüler', sStat);
            personChip('SGA', sgaStat);
            personChip('Schülervertretung', svStat);

            // Klassen-Gruppen
            if (clTotal === 0) {
                chip('bi-dash', 'Klassen: keine Einträge', 'muted', 'Keine Klassen eingetragen.');
            } else if (clMatched === clTotal) {
                chip('bi-check-circle-fill', 'Klassen: ' + clTotal + ' gematcht', 'ok', 'Alle Klassen-Gruppen in Microsoft 365 gefunden.');
            } else if (clUnchecked > 0) {
                chip('bi-question-circle', 'Klassen: ' + clUnchecked + '/' + clTotal + ' ungeprüft', 'warn',
                    clUnchecked + ' noch nicht geprüft, ' + clNotFound + ' nicht gefunden, ' + clMatched + ' gefunden.');
            } else {
                chip('bi-x-circle', 'Klassen: ' + clNotFound + '/' + clTotal + ' nicht gefunden', 'error',
                    clNotFound + ' Klassen-Gruppen fehlen in Microsoft 365, ' + clMatched + ' gefunden.');
            }

            // Fächer / ARGEs – reine Zahl (kein Gruppen-Match vorhanden)
            if (subjectRows.length > 0) {
                let subLinked = 0;
                subjectRows.forEach(function (r) {
                    const l = getSubjectCatalogLink(r.code);
                    if (l && normStr(l.graphGroupId)) subLinked++;
                });
                const subTitle =
                    subLinked > 0
                        ? subLinked + ' von ' + subjectRows.length + ' Fachgruppen mit Microsoft-365-Gruppe verknüpft.'
                        : subjectRows.length +
                          ' Fachgruppen eingetragen – „Aus Microsoft 365 übernehmen“ oder ARGE & Fachgruppen.';
                chip(
                    subLinked === subjectRows.length && subjectRows.length > 0 ? 'bi-check-circle-fill' : 'bi-info-circle',
                    subjectRows.length + ' Fachgruppen' + (subLinked ? ' (' + subLinked + ' M365)' : ''),
                    subLinked === subjectRows.length && subjectRows.length > 0 ? 'ok' : 'muted',
                    subTitle
                );
            }
            if (argeRows.length > 0) {
                let linkedN = 0;
                argeRows.forEach(function (r) {
                    const l = getArgeCatalogLink(r.code);
                    if (l && normStr(l.graphGroupId)) linkedN++;
                });
                const argeTitle =
                    linkedN > 0
                        ? linkedN + ' von ' + argeRows.length + ' ARGEs mit Microsoft-365-Gruppe verknüpft.'
                        : argeRows.length + ' ARGEs eingetragen – „Aus Microsoft 365 übernehmen“ oder Einrichtung für Verknüpfungen.';
                chip(
                    linkedN === argeRows.length && argeRows.length > 0 ? 'bi-check-circle-fill' : 'bi-info-circle',
                    argeRows.length + ' ARGEs' + (linkedN ? ' (' + linkedN + ' M365)' : ''),
                    linkedN === argeRows.length && argeRows.length > 0 ? 'ok' : 'muted',
                    argeTitle
                );
            }

            // DOM rendern
            grid.replaceChildren();
            chips.forEach(function (c) {
                const span = document.createElement('span');
                span.className = 'ts-status-chip ts-status-chip--' + c.kind;
                if (c.title) { span.title = c.title; span.setAttribute('aria-label', c.label + ': ' + c.title); }
                span.innerHTML = '<i class="bi ' + escapeHtml(c.icon) + '" aria-hidden="true"></i> ' + escapeHtml(c.label);
                grid.appendChild(span);
            });

            if (footer) {
                const now = new Date();
                footer.textContent = 'Zuletzt aktualisiert: ' + now.toLocaleTimeString('de-AT', { hour: '2-digit', minute: '2-digit' });
            }
        }

    function graphApi() {
        const api = window.ms365GraphUnifiedGroups;
        if (!api) throw new Error('Microsoft-Graph-Modul fehlt.');
        return api;
    }

    function randomTempPassword() {
        const u = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
        const l = 'abcdefghijkmnopqrstuvwxyz';
        const d = '23456789';
        const s = '!@#$%&*';
        function pick(set) {
            return set.charAt(Math.floor(Math.random() * set.length));
        }
        let pwd = pick(u) + pick(l) + pick(d) + pick(s);
        for (let i = 0; i < 12; i++) pwd += pick(u + l + d);
        return pwd;
    }

    function mailNicknameFromUpn(upn) {
        const local = String(upn || '').split('@')[0] || 'user';
        return graphApi().sanitizeMailNickname(local.replace(/[^a-zA-Z0-9]/g, '') || 'user');
    }

    function patchDirectoryMatchKeys(emailKeys, payload) {
        const patch = {};
        const seen = new Set();
        (emailKeys || []).forEach(function (k) {
            const em = normStr(k).toLowerCase();
            if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
            seen.add(em);
            patch[em] = payload;
        });
        if (!Object.keys(patch).length) return;
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: patch });
        }
    }

    function directoryMatchUserPayload(user, iso) {
        return {
            graphUserId: String(user && user.id ? user.id : '').trim(),
            displayName: String((user && (user.displayName || user.displayNameHint)) || '').trim(),
            userPrincipalName: String((user && (user.userPrincipalName || user.upn)) || '').trim(),
            notFound: false,
            checkedAt: iso || new Date().toISOString()
        };
    }

    async function createEntraUserViaGraph(token, upn, displayName) {
        const pwd = randomTempPassword();
        const mailNick = mailNicknameFromUpn(upn);
        const body = {
            accountEnabled: true,
            displayName: String(displayName).trim(),
            mailNickname: mailNick,
            userPrincipalName: String(upn).trim(),
            passwordProfile: {
                forceChangePasswordNextSignIn: true,
                password: pwd
            }
        };
        const created = await graphApi().graphJson('POST', '/users', token, body, undefined);
        const uid = String(created.id || '').trim();
        if (!uid) throw new Error('Keine Benutzer-ID von Graph erhalten.');
        return {
            id: uid,
            password: pwd,
            upn: String(upn).trim(),
            displayName: String(displayName).trim(),
            mailNickname: mailNick
        };
    }

    function getDirectoryMatchByEmail(emailRaw) {
        const em = normStr(emailRaw).toLowerCase();
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const map = setup && setup.directoryMatchByEmail ? setup.directoryMatchByEmail : {};
        return em ? map[em] || null : null;
    }

    async function verifyAdminDirectoryEmail(emailRaw) {
        const em = normStr(emailRaw).toLowerCase();
        if (!em || em.indexOf('@') === -1) {
            setSummary('Bitte zuerst eine gültige E-Mail eintragen.', 'warn');
            return false;
        }
        try {
            const token = await graphApi().getGraphToken();
            const u = await graphApi().resolveUserByEmail(token, em);
            const iso = new Date().toISOString();
            if (u && u.id) {
                patchDirectoryMatchKeys([em], directoryMatchUserPayload(u, iso));
                setSummary('Microsoft 365 gefunden: ' + (u.displayName || em), 'ok');
            } else {
                patchDirectoryMatchKeys([em], { notFound: true, checkedAt: iso });
                setSummary('Kein Entra-Benutzer für: ' + em, 'warn');
            }
            renderAdminUnifiedTableFromBundle();
            return !!(u && u.id);
        } catch (e) {
            setSummary('Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
            return false;
        }
    }

    async function runVerifyVerwaltungGraphBulk() {
        if (!btnVerifyVerwaltungGraph) return;
        await runVerifyGraphDirectoryRows(
            groupsToDisplayRows(getAdministrationGroups()),
            function (r) {
                return r && r.email;
            },
            'Verwaltung',
            btnVerifyVerwaltungGraph,
            function () {
                renderAdminUnifiedTableFromBundle();
            }
        );
    }

    function buildDirectoryMatchCell(tr, emailRaw, dirMatchMapOpt) {
        const tdMs = document.createElement('td');
        tdMs.style.fontSize = '0.88em';
        tdMs.style.lineHeight = '1.35';
        const em = normStr(emailRaw).toLowerCase();
        let m = null;
        if (em && em.indexOf('@') !== -1) {
            if (dirMatchMapOpt && typeof dirMatchMapOpt === 'object') {
                m = dirMatchMapOpt[em] || null;
            } else {
                m = getDirectoryMatchByEmail(em);
            }
        }
        if (!em || em.indexOf('@') === -1) {
            tdMs.style.color = 'var(--muted)';
            tdMs.textContent = '–';
            tdMs.title = 'E‑Mail nötig für Abgleich mit Microsoft Entra';
        } else if (m && m.graphUserId) {
            const gid = String(m.graphUserId);
            const short = gid.length > 14 ? gid.slice(0, 12) + '…' : gid;
            tdMs.innerHTML =
                '<span style="color:#0d8050;font-weight:700;">✓</span> <code style="font-size:0.82em;">' +
                escapeHtml(short) +
                '</code>';
            tdMs.title =
                (m.displayName ? m.displayName : '') +
                (m.userPrincipalName ? '\n' + m.userPrincipalName : '') +
                '\nObject-ID: ' +
                gid;
            tr.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
        } else if (m && m.notFound) {
            tdMs.innerHTML =
                '<span style="color:#856404;font-weight:700;">✗</span> <span style="color:var(--muted)">nicht gefunden</span>';
            tdMs.title = 'Kein Benutzer mit mail oder UPN gleich dieser E‑Mail';
        } else {
            tdMs.style.color = 'var(--muted)';
            tdMs.textContent = '–';
            tdMs.title = 'Noch nicht geprüft – Prüfen-Button in der Aktionsspalte';
        }
        return tdMs;
    }

    async function runVerifyGraphDirectoryRows(rows, getEmail, label, btn, onDone) {
        const updates = {};
        let found = 0;
        let missed = 0;
        let skipped = 0;
        const seen = new Set();
        const unique = [];
        const iso = new Date().toISOString();
        (rows || []).forEach(function (row) {
            const em = normStr(getEmail(row) || '').toLowerCase();
            if (!em || em.indexOf('@') === -1) {
                skipped++;
                return;
            }
            if (seen.has(em)) return;
            seen.add(em);
            unique.push(em);
        });
        try {
            if (btn) {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
            }
            if (!unique.length) {
                setSummary(label + ': keine gültigen E‑Mails zum Abgleich.', 'warn');
                return;
            }
            setSummary(
                label +
                    ': Microsoft-365-Abgleich läuft … (' +
                    unique.length +
                    ' Adresse' +
                    (unique.length === 1 ? '' : 'n') +
                    ')',
                'warn'
            );
            const api = graphApi();
            const token = await api.getGraphToken();
            let byEmail = null;
            if (typeof api.resolveUsersByEmailsBulk === 'function') {
                const bulk = await api.resolveUsersByEmailsBulk(token, unique, function (p) {
                    if (!p) return;
                    if (p.phase === 'directory') {
                        setSummary(
                            label +
                                ': Verzeichnis laden … ' +
                                (p.loaded || 0) +
                                ' Benutzer · Abgleich ' +
                                unique.length +
                                ' E‑Mails',
                            'warn'
                        );
                    } else if (p.phase === 'lookup') {
                        setSummary(
                            label + ': Abgleich ' + ((p.done || 0) + 1) + '/' + (p.total || unique.length) + ' …',
                            'warn'
                        );
                    }
                });
                byEmail = bulk && bulk.byEmail ? bulk.byEmail : new Map();
            } else {
                byEmail = new Map();
                for (let i = 0; i < unique.length; i++) {
                    const em = unique[i];
                    setSummary(label + ': Abgleich ' + (i + 1) + '/' + unique.length + ' …', 'warn');
                    try {
                        const u = await api.resolveUserByEmail(token, em);
                        byEmail.set(em, u && u.id ? u : null);
                    } catch {
                        byEmail.set(em, null);
                    }
                }
            }
            unique.forEach(function (em) {
                const u = byEmail.get(em);
                if (u && u.id) {
                    updates[em] = directoryMatchUserPayload(u, iso);
                    found++;
                } else {
                    updates[em] = { notFound: true, checkedAt: iso };
                    missed++;
                }
            });
            if (Object.keys(updates).length && window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: updates });
            }
            if (typeof onDone === 'function') onDone();
            renderStatusOverview();
            setSummary(
                label + ': ' + found + ' gefunden, ' + missed + ' nicht gefunden, ' + skipped + ' ohne gültige E‑Mail',
                missed ? 'warn' : 'ok'
            );
        } catch (e) {
            setSummary('Microsoft-365-Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

    function classMatchKey(row) {
        const code = normCode(row && row.code);
        const name = normCode(row && row.name);
        return code || name || '';
    }

    function getClassGroupMatchByKey(keyRaw) {
        const key = normCode(keyRaw);
        if (!key) return null;
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const map = setup && setup.classGroupMatchByKey ? setup.classGroupMatchByKey : {};
        return map[key] || null;
    }

    function patchClassGroupMatchByKey(keyRaw, payload) {
        const key = normCode(keyRaw);
        if (!key) return;
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
        window.ms365AppDataV2.patchSetup({ classGroupMatchByKey: { [key]: payload } });
    }

    function linkClassGroupMatch(row, group) {
        const key = classMatchKey(row);
        if (!key || !group || !group.id) return;
        patchClassGroupMatchByKey(key, {
            groupId: normStr(group.id),
            displayName: normStr(group.displayName),
            mailNickname: normStr(group.mailNickname),
            mail: normStr(group.mail),
            notFound: false,
            checkedAt: new Date().toISOString()
        });
    }

    function buildClassGroupMatchCell(tr, row) {
        const td = document.createElement('td');
        td.style.fontSize = '0.88em';
        td.style.lineHeight = '1.35';
        const key = classMatchKey(row);
        if (!key) {
            td.style.color = 'var(--muted)';
            td.textContent = '?';
            td.title = 'Abkürzung oder Klassenname nötig für den Abgleich';
            return td;
        }
        const m = getClassGroupMatchByKey(key);
        if (m && m.groupId) {
            const short = String(m.groupId).trim();
            const show = short.length > 14 ? short.slice(0, 12) + '…' : short;
            td.innerHTML =
                '<span style="color:#0d8050;font-weight:700;">✓</span> <code style="font-size:0.82em;">' +
                escapeHtml(show) +
                '</code>';
            td.title =
                (m.displayName ? String(m.displayName) : '') +
                (m.mailNickname ? '\nAlias: ' + String(m.mailNickname) : '') +
                (m.mail ? '\nMail: ' + String(m.mail) : '') +
                '\nGroup-ID: ' +
                short;
            tr.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
        } else if (m && m.notFound) {
            td.innerHTML =
                '<span style="color:#856404;font-weight:700;">✗</span> <span style="color:var(--muted)">nicht gefunden</span>';
            td.title = 'Keine passende Klassengruppe gefunden';
        } else {
            td.style.color = 'var(--muted)';
            td.textContent = '?';
            td.title = 'Noch nicht geprüft';
        }
        return td;
    }

    function pickClassGroupMatchFromSearch(groups, row) {
        const key = classMatchKey(row);
        const name = normStr(row && row.name);
        const keyLc = key.toLowerCase();
        const nameLc = name.toLowerCase();
        const sanitize =
            graphApi() && typeof graphApi().sanitizeMailNickname === 'function'
                ? graphApi().sanitizeMailNickname
                : function (v) {
                      return String(v || '')
                          .replace(/[^0-9a-zA-Z]/g, '')
                          .toLowerCase();
                  };
        const keyNick = sanitize(key);
        const list = Array.isArray(groups) ? groups : [];
        for (let i = 0; i < list.length; i++) {
            const g = list[i] || {};
            const dn = normStr(g.displayName).toLowerCase();
            const nick = normStr(g.mailNickname).toLowerCase();
            const mailLocal = normStr(g.mail).split('@')[0].toLowerCase();
            if (dn === keyLc || nick === keyNick || mailLocal === keyNick) return g;
            if (nameLc && dn === nameLc) return g;
        }
        return null;
    }

    function resolveClassGroupFromTenantInventory(row) {
        const inv = window.ms365TenantInventory;
        if (!inv || typeof inv.loadStructureState !== 'function' || typeof inv.readCache !== 'function') return null;
        const key = classMatchKey(row);
        if (!key) return null;
        const st = inv.loadStructureState();
        const rows = st && Array.isArray(st.rows) ? st.rows : [];
        const links = typeof inv.loadMatchLinks === 'function' ? inv.loadMatchLinks() || {} : {};
        const cls = rows.find(function (r) {
            return r && String(r.typ || '') === 'Klasse' && normCode(r.bezeichnung) === key;
        });
        if (!cls) return null;
        let gid = normStr(cls.tenantGroupId) || (links[String(cls.id)] && normStr(links[String(cls.id)].tenantGroupId)) || '';
        if (!gid && typeof inv.suggestGroupForUnit === 'function') gid = normStr(inv.suggestGroupForUnit(cls));
        if (!gid) return { notFound: true };
        const cache = inv.readCache() || { rows: [] };
        const hit = (cache.rows || []).find(function (g) {
            return normStr(g && g.id) === gid;
        });
        if (hit) {
            return {
                groupId: gid,
                displayName: normStr(hit.bezeichnung || hit.displayName),
                mailNickname: normStr(hit.alias || hit.mailNickname),
                mail: normStr(hit.mail),
                notFound: false,
                checkedAt: new Date().toISOString()
            };
        }
        return { groupId: gid, notFound: false, checkedAt: new Date().toISOString() };
    }

    async function verifyClassGroupForRow(row, opts) {
        const key = classMatchKey(row);
        if (!key) return { skipped: true };
        const invHit = resolveClassGroupFromTenantInventory(row);
        if (invHit) {
            patchClassGroupMatchByKey(key, invHit.groupId ? invHit : { notFound: true, checkedAt: new Date().toISOString() });
            return { found: !!invHit.groupId, skipped: false };
        }
        const token = opts && opts.token ? opts.token : await graphApi().getGraphToken();
        const queries = [key, normStr(row && row.name)].filter(Boolean);
        const seen = new Set();
        const groups = [];
        for (let i = 0; i < queries.length; i++) {
            const q = queries[i];
            let hits = [];
            if (typeof graphApi().searchUnifiedGroups === 'function') {
                hits = await graphApi().searchUnifiedGroups(token, q);
            } else {
                const data = await graphApi().graphJson(
                    'GET',
                    '/groups?$filter=' +
                        encodeURIComponent("groupTypes/any(c:c eq 'Unified') and startswith(displayName,'" + String(q).replace(/'/g, "''") + "')") +
                        '&$select=' +
                        encodeURIComponent('id,displayName,mail,mailNickname,groupTypes') +
                        '&$top=25',
                    token,
                    undefined
                );
                hits = Array.isArray(data && data.value) ? data.value : [];
            }
            hits.forEach(function (g) {
                const id = normStr(g && g.id);
                if (!id || seen.has(id)) return;
                seen.add(id);
                groups.push(g);
            });
        }
        const match = pickClassGroupMatchFromSearch(groups, row);
        if (match && match.id) {
            patchClassGroupMatchByKey(key, {
                groupId: normStr(match.id),
                displayName: normStr(match.displayName),
                mailNickname: normStr(match.mailNickname),
                mail: normStr(match.mail),
                notFound: false,
                checkedAt: new Date().toISOString()
            });
            return { found: true, skipped: false };
        }
        patchClassGroupMatchByKey(key, { notFound: true, checkedAt: new Date().toISOString() });
        return { found: false, skipped: false };
    }

    async function createAdminEntraUserInteractive(emailRaw, nameHint) {
        const em = normStr(emailRaw).toLowerCase();
        if (!em || em.indexOf('@') === -1) {
            setSummary('Bitte zuerst eine gültige E-Mail eintragen.', 'warn');
            return;
        }
        const dom = em.split('@')[1] || '';
        const upn = await dlgPrompt('Benutzer-Principalname (UPN), z. B. vorname.nachname@' + dom + ':', em, {
            title: 'Entra-Benutzer',
            inputLabel: 'UPN'
        });
        if (upn == null || !normStr(upn)) return;
        const displayName = await dlgPrompt(
            'Anzeigename in Microsoft 365:',
            normStr(nameHint) || normStr(String(upn).split('@')[0]),
            { title: 'Entra-Benutzer', inputLabel: 'Anzeigename' }
        );
        if (displayName == null || !normStr(displayName)) return;
        try {
            const tokenProbe = await graphApi().getGraphToken();
            const existing = await graphApi().resolveUserByEmail(tokenProbe, em);
            if (existing && existing.id) {
                patchDirectoryMatchKeys([em], directoryMatchUserPayload(existing));
                renderAdminUnifiedTableFromBundle();
                setSummary('Unter dieser E-Mail existiert bereits ein Entra-Benutzer.', 'ok');
                return;
            }
        } catch {
            // weiter mit Anlage
        }
        const mailNick = mailNicknameFromUpn(upn);
        const ok = await dlgConfirm(
            'Benutzer in Entra ID anlegen?\n\nAnzeigename: ' +
                displayName +
                '\nUPN: ' +
                upn +
                '\nMail-Nickname: ' +
                mailNick +
                '\n\nEs wird ein temporäres Kennwort gesetzt (Wechsel beim ersten Anmelden).',
            { title: 'Entra-Benutzer anlegen' }
        );
        if (!ok) return;
        try {
            const token = await graphApi().getGraphToken();
            const created = await createEntraUserViaGraph(token, upn, displayName);
            const iso = new Date().toISOString();
            patchDirectoryMatchKeys(
                [em, created.upn],
                directoryMatchUserPayload(
                    { id: created.id, displayName: created.displayName, userPrincipalName: created.upn },
                    iso
                )
            );
            renderAdminUnifiedTableFromBundle();
            await dlgAlert(
                'Benutzer angelegt.\n\nEinmaliges Kennwort:\n' +
                    created.password +
                    '\n\n(Bitte sicher übergeben / in Entra ändern.)',
                { title: 'Kennwort notieren', okText: 'Verstanden' }
            );
            setSummary('Entra-Benutzer angelegt: ' + created.displayName, 'ok');
        } catch (e) {
            setSummary('Benutzer anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
        }
    }

        function subjectsToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code)};${normStr(x.name || '')}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getSubjectsFromTextarea() {
            return typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
        }

        function setSubjectsTextareaFromRows(rows) {
            if (!taSubjects) return;
            taSubjects.value = subjectsToLines(rows);
        }

        function renderSubjectsTableFromTextarea() {
            if (!subjectsTbody) return;
            updateTableSortIndicators(subjectsTable, subjectsSortState);
            const rows = getSubjectsFromTextarea();
            subjectsTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen, aus Microsoft 365 übernehmen oder „+ Zeile“.';
                tr.appendChild(td);
                subjectsTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getSubjectsFromTextarea();
                        if (!all[idx]) return renderSubjectsTableFromTextarea();
                        const prev = all[idx].code;
                        const newCode = meta && meta.cancelled ? prev : normCode(next);
                        if (prev && newCode && normCode(prev) !== normCode(newCode)) {
                            const api = window.ms365AppDataV2;
                            if (api && typeof api.renameCatalogLink === 'function') {
                                try {
                                    api.renameCatalogLink('subject', prev, newCode);
                                } catch {
                                    /* Verknüpfung bleibt am alten Kürzel */
                                }
                            }
                        }
                        all[idx].code = newCode;
                        setSubjectsTextareaFromRows(all);
                        renderSubjectsTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getSubjectsFromTextarea();
                        if (!all[idx]) return renderSubjectsTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setSubjectsTextareaFromRows(all);
                        renderSubjectsTableFromTextarea();
                        scheduleAutoSave();
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
                    const all = getSubjectsFromTextarea();
                    all.splice(idx, 1);
                    setSubjectsTextareaFromRows(all);
                    renderSubjectsTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, buildSubjectM365Cell(row), tdAction);
                subjectsTbody.appendChild(tr);
            });
        }

        function teachersToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code)};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function adminToLines(rows) {
            return (rows || [])
                .map((x) => `${normStr(x.role || '')};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function adminRolesToLines(rows) {
            return (rows || [])
                .map((x) => `${normCode(x.code || '')};${normStr(x.name || '')}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getAdministrationGroups() {
            if (taAdminBundle && typeof window.ms365TenantSettingsParseAdminGroupsLines === 'function') {
                return window.ms365TenantSettingsParseAdminGroupsLines(taAdminBundle.value);
            }
            const roles =
                typeof parseLinesToAdminRoles === 'function'
                    ? parseLinesToAdminRoles(taAdminRoles ? taAdminRoles.value : '')
                    : [];
            const admin =
                typeof parseLinesToAdmin === 'function' ? parseLinesToAdmin(taAdmin ? taAdmin.value : '') : [];
            if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                return window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, admin);
            }
            return [];
        }

        function setAdministrationGroups(groups) {
            if (taAdminBundle && typeof window.ms365TenantSettingsAdminGroupsToLines === 'function') {
                taAdminBundle.value = window.ms365TenantSettingsAdminGroupsToLines(groups);
                return;
            }
            if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups !== 'function') return;
            const roles = (groups || []).map(function (group) {
                return { code: normCode(group && group.code), name: normStr(group && group.name) };
            });
            const admin = [];
            (groups || []).forEach(function (group) {
                const roleName = normStr(group && group.name);
                (Array.isArray(group && group.people) ? group.people : []).forEach(function (person) {
                    admin.push({
                        role: roleName,
                        name: normStr(person && person.name),
                        email: normStr(person && person.email).toLowerCase(),
                        defaultKey: normStr(person && person.defaultKey)
                    });
                });
            });
            if (taAdminRoles) taAdminRoles.value = adminRolesToLines(roles);
            if (taAdmin) taAdmin.value = adminToLines(admin);
        }

        function groupsToDisplayRows(groups) {
            const rows = [];
            (Array.isArray(groups) ? groups : []).forEach(function (group) {
                const code = normStr(group && group.code);
                const name = normStr(group && group.name);
                const people = Array.isArray(group && group.people) ? group.people : [];
                if (!people.length) {
                    rows.push({ code: code, name: name, personName: '', email: '' });
                    return;
                }
                people.forEach(function (person) {
                    rows.push({
                        code: code,
                        name: name,
                        personName: normStr(person && person.name),
                        email: normStr(person && person.email).toLowerCase()
                    });
                });
            });
            return rows;
        }

        function displayRowsToGroups(rows) {
            const groupMap = new Map();
            const order = [];
            (Array.isArray(rows) ? rows : []).forEach(function (row) {
                const name = normStr(row && row.name);
                let code = normCode(row && row.code);
                if (!code && name && typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function') {
                    code = window.ms365TenantSettingsAdminRoleCodeFromName(name);
                }
                const personName = normStr(row && row.personName);
                const email = normStr(row && row.email).toLowerCase();
                if (!name && !code && !personName && !email) return;
                const key = (name || code).toLowerCase();
                if (!groupMap.has(key)) {
                    groupMap.set(key, { code: code, name: name || code, people: [] });
                    order.push(key);
                }
                const group = groupMap.get(key);
                if (code) group.code = code;
                if (name) group.name = name;
                if (personName || email) {
                    group.people.push({ name: personName, email: email });
                }
            });
            return order.map(function (key) {
                return groupMap.get(key);
            });
        }

        function adminGroupsFromSettings(settings) {
            if (
                settings &&
                Array.isArray(settings.administration) &&
                settings.administration.some(function (entry) {
                    return entry && Array.isArray(entry.people);
                })
            ) {
                return settings.administration;
            }
            if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                return window.ms365TenantSettingsAdminRolesAndAdminToGroups(
                    settings && settings.adminRoles ? settings.adminRoles : [],
                    settings && settings.admin ? settings.admin : []
                );
            }
            return [];
        }

        function getAdministrationEntries() {
            return getAdministrationGroups();
        }

        function getAdminFromTextarea() {
            const out = [];
            getAdministrationGroups().forEach(function (group) {
                const roleName = normStr(group && group.name);
                (Array.isArray(group && group.people) ? group.people : []).forEach(function (person) {
                    const row = {
                        role: roleName,
                        name: normStr(person && person.name),
                        email: normStr(person && person.email).toLowerCase()
                    };
                    if (person && person.defaultKey) row.defaultKey = normStr(person.defaultKey);
                    out.push(row);
                });
            });
            return out;
        }

        function setAdminTextareaFromRows(rows) {
            const roles = getAdminRolesFromTextarea();
            if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                setAdministrationGroups(window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, rows));
                return;
            }
            if (!taAdmin) return;
            taAdmin.value = adminToLines(rows);
        }

        function getAdminRolesFromTextarea() {
            return getAdministrationGroups().map(function (group) {
                return { code: normCode(group && group.code), name: normStr(group && group.name) };
            });
        }

        function setAdminRolesTextareaFromRows(rows) {
            const people = getAdminFromTextarea();
            if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                setAdministrationGroups(window.ms365TenantSettingsAdminRolesAndAdminToGroups(rows, people));
                return;
            }
            if (taAdminRoles) taAdminRoles.value = adminRolesToLines(rows);
        }

        function renderAdminUnifiedTableFromBundle() {
            if (!adminUnifiedTbody) return;
            updateTableSortIndicators(adminUnifiedTable, adminUnifiedSortState);
            const displayRows = groupsToDisplayRows(getAdministrationGroups());
            adminUnifiedTbody.replaceChildren();
            if (!displayRows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 6;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Eintrag“.';
                tr.appendChild(td);
                adminUnifiedTbody.appendChild(tr);
                return;
            }

            displayRows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdLabel = document.createElement('td');
                tdLabel.textContent = row.name || '';
                tdLabel.title = 'Doppelklick zum Bearbeiten';
                tdLabel.addEventListener('dblclick', () => {
                    startCellEdit(tdLabel, row.name, (next, meta) => {
                        const all = groupsToDisplayRows(getAdministrationGroups());
                        if (!all[idx]) return renderAdminUnifiedTableFromBundle();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        if (all[idx].code && prev && all[idx].name !== prev) {
                            all[idx].code =
                                typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function'
                                    ? window.ms365TenantSettingsAdminRoleCodeFromName(all[idx].name)
                                    : all[idx].code;
                        }
                        setAdministrationGroups(displayRowsToGroups(all));
                        renderAdminUnifiedTableFromBundle();
                        scheduleAutoSave();
                    });
                });

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = groupsToDisplayRows(getAdministrationGroups());
                        if (!all[idx]) return renderAdminUnifiedTableFromBundle();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setAdministrationGroups(displayRowsToGroups(all));
                        renderAdminUnifiedTableFromBundle();
                        scheduleAutoSave();
                    });
                });

                const tdPerson = document.createElement('td');
                tdPerson.textContent = row.personName || '';
                tdPerson.title = 'Doppelklick zum Bearbeiten';
                tdPerson.addEventListener('dblclick', () => {
                    startCellEdit(tdPerson, row.personName, (next, meta) => {
                        const all = groupsToDisplayRows(getAdministrationGroups());
                        if (!all[idx]) return renderAdminUnifiedTableFromBundle();
                        const prev = all[idx].personName;
                        all[idx].personName = meta && meta.cancelled ? prev : normStr(next);
                        setAdministrationGroups(displayRowsToGroups(all));
                        renderAdminUnifiedTableFromBundle();
                        scheduleAutoSave();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = groupsToDisplayRows(getAdministrationGroups());
                        if (!all[idx]) return renderAdminUnifiedTableFromBundle();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setAdministrationGroups(displayRowsToGroups(all));
                        renderAdminUnifiedTableFromBundle();
                        scheduleAutoSave();
                    });
                });

                const tdMs = document.createElement('td');
                tdMs.style.fontSize = '0.88em';
                tdMs.style.lineHeight = '1.35';
                {
                    const em = row.email ? row.email.trim().toLowerCase() : '';
                    const m = em && em.indexOf('@') !== -1 ? getDirectoryMatchByEmail(em) : null;
                    if (!em || em.indexOf('@') === -1) {
                        tdMs.style.color = 'var(--muted)';
                        tdMs.textContent = '–';
                        tdMs.title = 'E‑Mail nötig für Abgleich mit Microsoft Entra';
                    } else if (m && m.graphUserId) {
                        const gid = String(m.graphUserId);
                        const short = gid.length > 14 ? gid.slice(0, 12) + '…' : gid;
                        tdMs.innerHTML =
                            '<span style="color:#0d8050;font-weight:700;">✓</span> <code style="font-size:0.82em;">' +
                            escapeHtml(short) + '</code>';
                        tdMs.title =
                            (m.displayName ? m.displayName : '') +
                            (m.userPrincipalName ? '\n' + m.userPrincipalName : '') +
                            '\nObject-ID: ' + gid;
                        tr.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
                    } else if (m && m.notFound) {
                        tdMs.innerHTML =
                            '<span style="color:#856404;font-weight:700;">✗</span> <span style="color:var(--muted)">nicht gefunden</span>';
                        tdMs.title = 'Kein Benutzer mit mail oder UPN gleich dieser E‑Mail';
                    } else {
                        tdMs.style.color = 'var(--muted)';
                        tdMs.textContent = '–';
                        tdMs.title = 'Noch nicht geprüft – Prüfen-Button in der Aktionsspalte';
                    }
                }

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                tdAction.style.whiteSpace = 'nowrap';
                tdAction.style.display = 'flex';
                tdAction.style.gap = '6px';
                tdAction.style.alignItems = 'center';
                const dir = getDirectoryMatchByEmail(row.email);
                if (row.email && row.personName) {
                    const btnCheck = document.createElement('button');
                    btnCheck.type = 'button';
                    btnCheck.className = 'mini-btn';
                    btnCheck.style.background = 'var(--brand1)';
                    btnCheck.title = 'Diese E‑Mail in Microsoft Entra prüfen';
                    btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCheck.disabled = true;
                    } else {
                        btnCheck.addEventListener('click', async () => {
                            btnCheck.disabled = true;
                            try {
                                await verifyAdminDirectoryEmail(row.email);
                            } finally {
                                renderAdminUnifiedTableFromBundle();
                            }
                        });
                    }
                    tdAction.appendChild(btnCheck);

                    const btnCreate = document.createElement('button');
                    btnCreate.type = 'button';
                    btnCreate.className = 'mini-btn';
                    btnCreate.style.background = '#11cdef';
                    btnCreate.title = 'Neuen Benutzer in Microsoft Entra ID anlegen (User.ReadWrite.All)';
                    btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCreate.disabled = true;
                    } else {
                        btnCreate.addEventListener('click', async () => {
                            btnCreate.disabled = true;
                            try {
                                await createAdminEntraUserInteractive(row.email, row.personName);
                            } finally {
                                renderAdminUnifiedTableFromBundle();
                            }
                        });
                    }
                    tdAction.appendChild(btnCreate);
                }
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = groupsToDisplayRows(getAdministrationGroups());
                    all.splice(idx, 1);
                    setAdministrationGroups(displayRowsToGroups(all));
                    renderAdminUnifiedTableFromBundle();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdLabel, tdCode, tdPerson, tdEmail, tdMs, tdAction);
                adminUnifiedTbody.appendChild(tr);
            });
        }

        function getTeachersFromTextarea() {
            return typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
        }

        function setTeachersTextareaFromRows(rows) {
            if (!taTeachers) return;
            taTeachers.value = teachersToLines(rows);
        }

        function sortValueForTenantRow(row, key) {
            if (!row) return '';
            if (key === 'subjects') {
                return (Array.isArray(row.subjects) ? row.subjects : [])
                    .map((s) => normCode(s))
                    .filter(Boolean)
                    .join(',');
            }
            if (key === 'code' || key === 'year') return normCode(String(row[key] || ''));
            if (key === 'email' || key === 'headEmail') return normStr(row[key]).toLowerCase();
            if (key === 'scope') return normStr(row.scope).toLowerCase();
            if (key === 'klasse') return normStr(row.klasse).toLowerCase();
            if (key === 'personName') return normStr(row.personName);
            return normStr(row[key]);
        }

        function sortTenantRows(rows, key, dir) {
            return (rows || [])
                .slice()
                .sort((a, b) => {
                    const av = sortValueForTenantRow(a, key);
                    const bv = sortValueForTenantRow(b, key);
                    return av.localeCompare(bv, 'de', { sensitivity: 'base', numeric: true }) * dir;
                });
        }

        function updateTableSortIndicators(table, sortState) {
            if (!table) return;
            table.querySelectorAll('th.is-sortable').forEach((th) => {
                const label = th.dataset.label || th.textContent.replace(/\s*[▲▼]\s*$/, '').trim();
                th.dataset.label = label;
                th.textContent = label;
                th.setAttribute('aria-sort', 'none');
                if (sortState.key && th.dataset.sortKey === sortState.key) {
                    const ind = document.createElement('span');
                    ind.className = 'teachers-sort-ind';
                    ind.setAttribute('aria-hidden', 'true');
                    ind.textContent = sortState.dir === -1 ? '▼' : '▲';
                    th.appendChild(ind);
                    th.setAttribute('aria-sort', sortState.dir === -1 ? 'descending' : 'ascending');
                }
            });
        }

        function toggleTenantSortState(sortState, key) {
            if (sortState.key === key) sortState.dir = sortState.dir === 1 ? -1 : 1;
            else {
                sortState.key = key;
                sortState.dir = 1;
            }
        }

        function bindTenantTableSort(table, sortState, allowedKeys, onSortKey) {
            if (!table || table.dataset.tenantSortBound) return;
            table.dataset.tenantSortBound = '1';
            table.querySelectorAll('th.is-sortable').forEach((th) => {
                const key = th.dataset.sortKey;
                if (!key || (allowedKeys && allowedKeys.indexOf(key) < 0)) return;
                th.addEventListener('click', () => onSortKey(key));
                th.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        onSortKey(key);
                    }
                });
            });
        }

        function updateTeachersSortIndicators() {
            updateTableSortIndicators(teachersTable, teachersSortState);
        }

        function applyTeachersSort(key) {
            if (!key || (key !== 'code' && key !== 'name' && key !== 'email')) return;
            if (teachersSortState.key === key) {
                teachersSortState.dir = teachersSortState.dir === 1 ? -1 : 1;
            } else {
                teachersSortState.key = key;
                teachersSortState.dir = 1;
            }
            const api = window.ms365TeacherListImport;
            const all = getTeachersFromTextarea();
            const sorted =
                api && typeof api.sortRows === 'function'
                    ? api.sortRows(all, teachersSortState.key, teachersSortState.dir)
                    : all.slice().sort((a, b) => {
                          const av = String((a && a[key]) || '');
                          const bv = String((b && b[key]) || '');
                          return av.localeCompare(bv, 'de', { sensitivity: 'base' }) * teachersSortState.dir;
                      });
            setTeachersTextareaFromRows(sorted);
            renderTeachersTableFromTextarea();
            scheduleAutoSave();
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
            updateTeachersSortIndicators();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 5;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen, aus Microsoft 365 einlesen oder „+ Zeile“.';
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

                const tdMs = buildDirectoryMatchCell(tr, row.email);

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                tdAction.style.whiteSpace = 'nowrap';
                tdAction.style.display = 'flex';
                tdAction.style.gap = '6px';
                tdAction.style.alignItems = 'center';
                const dir = getDirectoryMatchByEmail(row.email);
                if (row.email && row.name) {
                    const btnCheck = document.createElement('button');
                    btnCheck.type = 'button';
                    btnCheck.className = 'mini-btn';
                    btnCheck.style.background = 'var(--brand1)';
                    btnCheck.title = 'Diese E‑Mail in Microsoft Entra prüfen';
                    btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCheck.disabled = true;
                    } else {
                        btnCheck.addEventListener('click', async () => {
                            btnCheck.disabled = true;
                            try {
                                await verifyAdminDirectoryEmail(row.email);
                            } finally {
                                renderTeachersTableFromTextarea();
                            }
                        });
                    }
                    tdAction.appendChild(btnCheck);

                    const btnCreate = document.createElement('button');
                    btnCreate.type = 'button';
                    btnCreate.className = 'mini-btn';
                    btnCreate.style.background = '#11cdef';
                    btnCreate.title = 'Neuen Benutzer in Microsoft Entra ID anlegen (User.ReadWrite.All)';
                    btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCreate.disabled = true;
                    } else {
                        btnCreate.addEventListener('click', async () => {
                            btnCreate.disabled = true;
                            try {
                                await createAdminEntraUserInteractive(row.email, row.name);
                            } finally {
                                renderTeachersTableFromTextarea();
                            }
                        });
                    }
                    tdAction.appendChild(btnCreate);
                }
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

                tr.append(tdCode, tdName, tdEmail, tdMs, tdAction);
                teachersTbody.appendChild(tr);
            });
        }

        function renderAdminRolesTableFromTextarea() {
            if (adminUnifiedTbody) {
                renderAdminUnifiedTableFromBundle();
                return;
            }
            if (!adminRolesTbody) return;
            const rows = getAdminRolesFromTextarea();
            adminRolesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 3;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Rollen – oben einfügen, „+ Rolle“ oder Standardrollen.';
                tr.appendChild(td);
                adminRolesTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getAdminRolesFromTextarea();
                        if (!all[idx]) return renderAdminRolesTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setAdminRolesTextareaFromRows(all);
                        renderAdminRolesTableFromTextarea();
                        renderAdminTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Umbenennen (Personen werden mitgezogen)';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getAdminRolesFromTextarea();
                        if (!all[idx]) return renderAdminRolesTableFromTextarea();
                        const prev = all[idx].name;
                        const nextName = meta && meta.cancelled ? prev : normStr(next);
                        if (nextName && nextName !== prev && typeof window.ms365TenantSettingsRenameAdminRole === 'function') {
                            const renamed = window.ms365TenantSettingsRenameAdminRole(all, getAdminFromTextarea(), prev, nextName);
                            setAdminRolesTextareaFromRows(renamed.roles);
                            setAdminTextareaFromRows(renamed.admin);
                        } else {
                            all[idx].name = nextName;
                            setAdminRolesTextareaFromRows(all);
                        }
                        renderAdminRolesTableFromTextarea();
                        renderAdminTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Rolle löschen';
                btnDel.addEventListener('click', () => {
                    const people = getAdminFromTextarea().filter((p) => {
                        if (typeof window.ms365TenantSettingsPersonMatchesAdminRole === 'function') {
                            return window.ms365TenantSettingsPersonMatchesAdminRole(p, row);
                        }
                        return normStr(p.role).toLowerCase() === normStr(row.name).toLowerCase();
                    });
                    if (people.length) {
                        window.alert(
                            'Diese Rolle ist noch ' +
                                people.length +
                                ' Person(en) zugeordnet. Bitte zuerst umbenennen oder die Personen entfernen.'
                        );
                        return;
                    }
                    const all = getAdminRolesFromTextarea();
                    all.splice(idx, 1);
                    setAdminRolesTextareaFromRows(all);
                    renderAdminRolesTableFromTextarea();
                    renderAdminTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdName, tdAction);
                adminRolesTbody.appendChild(tr);
            });
        }

        function renderAdminTableFromTextarea() {
            if (adminUnifiedTbody) {
                renderAdminUnifiedTableFromBundle();
                return;
            }
            if (!adminTbody) return;
            const rows = getAdminFromTextarea();
            const roleCatalog = getAdminRolesFromTextarea();
            adminTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Person“.';
                tr.appendChild(td);
                adminTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdRole = document.createElement('td');
                const sel = document.createElement('select');
                sel.style.width = '100%';
                sel.style.font = 'inherit';
                const names = [];
                const seenN = new Set();
                roleCatalog.forEach((r) => {
                    const n = normStr(r.name || r.code);
                    if (!n) return;
                    const k = n.toLowerCase();
                    if (seenN.has(k)) return;
                    seenN.add(k);
                    names.push(n);
                });
                const current = normStr(row.role);
                if (current && !seenN.has(current.toLowerCase())) names.unshift(current);
                const optEmpty = document.createElement('option');
                optEmpty.value = '';
                optEmpty.textContent = '— Rolle —';
                sel.appendChild(optEmpty);
                names.forEach((n) => {
                    const opt = document.createElement('option');
                    opt.value = n;
                    opt.textContent = n;
                    if (current && n.toLowerCase() === current.toLowerCase()) opt.selected = true;
                    sel.appendChild(opt);
                });
                sel.addEventListener('change', () => {
                    const all = getAdminFromTextarea();
                    if (!all[idx]) return renderAdminTableFromTextarea();
                    all[idx].role = normStr(sel.value);
                    setAdminTextareaFromRows(all);
                    renderAdminTableFromTextarea();
                    scheduleAutoSave();
                });
                tdRole.appendChild(sel);

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getAdminFromTextarea();
                        if (!all[idx]) return renderAdminTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setAdminTextareaFromRows(all);
                        renderAdminTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getAdminFromTextarea();
                        if (!all[idx]) return renderAdminTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setAdminTextareaFromRows(all);
                        renderAdminTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Person löschen';
                btnDel.addEventListener('click', () => {
                    const all = getAdminFromTextarea();
                    all.splice(idx, 1);
                    setAdminTextareaFromRows(all);
                    renderAdminTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdRole, tdName, tdEmail, tdAction);
                adminTbody.appendChild(tr);
            });
        }

        function sgaToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const scope =
                        x.scope === 'teacher'
                            ? 'Lehrer'
                            : x.scope === 'student'
                              ? 'Schueler'
                              : x.scope === 'external'
                                ? 'Extern'
                                : '';
                    return `${scope};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim();
                })
                .filter(Boolean)
                .join('\n');
        }

        function getSgaFromTextarea() {
            return typeof parseLinesToSga === 'function' ? parseLinesToSga(taSga ? taSga.value : '') : [];
        }

        function setSgaTextareaFromRows(rows) {
            if (!taSga) return;
            taSga.value = sgaToLines(rows);
        }

        function renderSgaTableFromTextarea() {
            if (!sgaTbody) return;
            updateTableSortIndicators(sgaTable, sgaSortState);
            const rows = getSgaFromTextarea();
            sgaTbody.replaceChildren();
            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 5;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine SGA-Mitglieder – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                sgaTbody.appendChild(tr);
                return;
            }
            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');
                const tdScope = document.createElement('td');
                const sel = document.createElement('select');
                sel.style.width = '100%';
                sel.style.font = 'inherit';
                [
                    { value: '', label: '— Gruppe —' },
                    { value: 'teacher', label: 'Lehrer' },
                    { value: 'student', label: 'Schüler' },
                    { value: 'external', label: 'Extern' }
                ].forEach((entry) => {
                    const opt = document.createElement('option');
                    opt.value = entry.value;
                    opt.textContent = entry.label;
                    if ((row.scope || '') === entry.value) opt.selected = true;
                    sel.appendChild(opt);
                });
                sel.addEventListener('change', () => {
                    const all = getSgaFromTextarea();
                    if (!all[idx]) return renderSgaTableFromTextarea();
                    all[idx].scope = normStr(sel.value);
                    setSgaTextareaFromRows(all);
                    renderSgaTableFromTextarea();
                    scheduleAutoSave();
                });
                tdScope.appendChild(sel);

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getSgaFromTextarea();
                        if (!all[idx]) return renderSgaTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setSgaTextareaFromRows(all);
                        renderSgaTableFromTextarea();
                        scheduleAutoSave();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getSgaFromTextarea();
                        if (!all[idx]) return renderSgaTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setSgaTextareaFromRows(all);
                        renderSgaTableFromTextarea();
                        scheduleAutoSave();
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
                    const all = getSgaFromTextarea();
                    all.splice(idx, 1);
                    setSgaTextareaFromRows(all);
                    renderSgaTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                const tdMs = buildDirectoryMatchCell(tr, row.email);

                // Per-row check and create buttons
                const btnCheck = document.createElement('button');
                btnCheck.type = 'button';
                btnCheck.className = 'mini-btn';
                btnCheck.title = 'In Microsoft Entra prüfen';
                btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                btnCheck.addEventListener('click', async () => {
                    const em = normStr(row.email || '').toLowerCase();
                    if (!em || em.indexOf('@') === -1) return setSummary('Keine gültige E-Mail für den Abgleich.', 'warn');
                    try {
                        const token = await graphApi().getGraphToken();
                        const u = await graphApi().resolveUserByEmail(token, em);
                        const iso = new Date().toISOString();
                        const setup = window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function' ? window.ms365AppDataV2.getSetup() : {};
                        const updates = Object.assign({}, (setup && setup.directoryMatchByEmail) || {});
                        if (u && u.id) {
                            updates[em] = directoryMatchUserPayload(u, iso);
                            setSummary(row.name + ': gefunden (' + (u.displayName || em) + ')', 'ok');
                        } else {
                            updates[em] = { notFound: true, checkedAt: iso };
                            setSummary(row.name + ': nicht in Microsoft Entra gefunden.', 'warn');
                        }
                        window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: updates });
                        renderSgaTableFromTextarea();
                    } catch (e) {
                        setSummary('Fehler beim Prüfen: ' + (e && e.message ? e.message : String(e)), 'warn');
                    }
                });
                tdAction.appendChild(btnCheck);

                const btnCreate = document.createElement('button');
                btnCreate.type = 'button';
                btnCreate.className = 'mini-btn';
                btnCreate.title = 'Benutzer in Microsoft Entra anlegen';
                btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
                btnCreate.addEventListener('click', () => {
                    if (typeof window.ms365TenantSettingsCreateUser === 'function') {
                        window.ms365TenantSettingsCreateUser(row, () => renderSgaTableFromTextarea());
                    } else {
                        setSummary('Anlegen nicht verfügbar – bitte in der geführten Einrichtung nutzen.', 'warn');
                    }
                });
                tdAction.appendChild(btnCreate);

                tr.append(tdScope, tdName, tdEmail, tdMs, tdAction);
                sgaTbody.appendChild(tr);
            });
        }

        function studentsToLines(rows) {
            return (rows || [])
                .map((x) => {
                    const base = `${normStr(x.klasse || '')};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`;
                    const pairs = Array.isArray(x.parentPairs) ? x.parentPairs : [];
                    if (!pairs.length) return base.trim();
                    const extra = pairs
                        .map((p) => `${normStr(p.name || '')};${normStr(p.email || '').toLowerCase()}`)
                        .join(';');
                    return `${base};${extra}`.trim();
                })
                .filter(Boolean)
                .join('\n');
        }

        function getStudentsFromTextarea() {
            return typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
        }

        function studentCouncilToLines(rows) {
            return (rows || [])
                .map((x) => `${normStr(x.klasse || '')};${normStr(x.name || '')};${normStr(x.email || '').toLowerCase()}`.trim())
                .filter(Boolean)
                .join('\n');
        }

        function getStudentCouncilFromTextarea() {
            return typeof parseLinesToStudentCouncil === 'function'
                ? parseLinesToStudentCouncil(taStudentCouncil ? taStudentCouncil.value : '')
                : [];
        }

        function setStudentCouncilTextareaFromRows(rows) {
            if (!taStudentCouncil) return;
            taStudentCouncil.value = studentCouncilToLines(rows);
        }

        function renderStudentCouncilTableFromTextarea() {
            if (!studentCouncilTbody) return;
            updateTableSortIndicators(studentCouncilTable, studentCouncilSortState);
            const rows = getStudentCouncilFromTextarea();
            studentCouncilTbody.replaceChildren();
            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 4;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                studentCouncilTbody.appendChild(tr);
                return;
            }
            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');
                const tdKlasse = document.createElement('td');
                tdKlasse.textContent = row.klasse || '';
                tdKlasse.title = 'Doppelklick zum Bearbeiten';
                tdKlasse.addEventListener('dblclick', () => {
                    startCellEdit(tdKlasse, row.klasse, (next, meta) => {
                        const all = getStudentCouncilFromTextarea();
                        if (!all[idx]) return renderStudentCouncilTableFromTextarea();
                        const prev = all[idx].klasse;
                        all[idx].klasse = meta && meta.cancelled ? prev : normStr(next);
                        setStudentCouncilTextareaFromRows(all);
                        renderStudentCouncilTableFromTextarea();
                        scheduleAutoSave();
                    });
                });
                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getStudentCouncilFromTextarea();
                        if (!all[idx]) return renderStudentCouncilTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setStudentCouncilTextareaFromRows(all);
                        renderStudentCouncilTableFromTextarea();
                        scheduleAutoSave();
                    });
                });
                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.email, (next, meta) => {
                        const all = getStudentCouncilFromTextarea();
                        if (!all[idx]) return renderStudentCouncilTableFromTextarea();
                        const prev = all[idx].email;
                        all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setStudentCouncilTextareaFromRows(all);
                        renderStudentCouncilTableFromTextarea();
                        scheduleAutoSave();
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
                    const all = getStudentCouncilFromTextarea();
                    all.splice(idx, 1);
                    setStudentCouncilTextareaFromRows(all);
                    renderStudentCouncilTableFromTextarea();
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);
                const tdMs = buildDirectoryMatchCell(tr, row.email);

                const btnCheck = document.createElement('button');
                btnCheck.type = 'button';
                btnCheck.className = 'mini-btn';
                btnCheck.title = 'In Microsoft Entra prüfen';
                btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                btnCheck.addEventListener('click', async () => {
                    const em = normStr(row.email || '').toLowerCase();
                    if (!em || em.indexOf('@') === -1) return setSummary('Keine gültige E-Mail für den Abgleich.', 'warn');
                    try {
                        const token = await graphApi().getGraphToken();
                        const u = await graphApi().resolveUserByEmail(token, em);
                        const iso = new Date().toISOString();
                        const setup = window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function' ? window.ms365AppDataV2.getSetup() : {};
                        const updates = Object.assign({}, (setup && setup.directoryMatchByEmail) || {});
                        if (u && u.id) {
                            updates[em] = directoryMatchUserPayload(u, iso);
                            setSummary(row.name + ': gefunden (' + (u.displayName || em) + ')', 'ok');
                        } else {
                            updates[em] = { notFound: true, checkedAt: iso };
                            setSummary(row.name + ': nicht in Microsoft Entra gefunden.', 'warn');
                        }
                        window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: updates });
                        renderStudentCouncilTableFromTextarea();
                    } catch (e) {
                        setSummary('Fehler beim Prüfen: ' + (e && e.message ? e.message : String(e)), 'warn');
                    }
                });
                tdAction.appendChild(btnCheck);

                const btnCreate = document.createElement('button');
                btnCreate.type = 'button';
                btnCreate.className = 'mini-btn';
                btnCreate.title = 'Benutzer in Microsoft Entra anlegen';
                btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
                btnCreate.addEventListener('click', () => {
                    if (typeof window.ms365TenantSettingsCreateUser === 'function') {
                        window.ms365TenantSettingsCreateUser(row, () => renderStudentCouncilTableFromTextarea());
                    } else {
                        setSummary('Anlegen nicht verfügbar – bitte in der geführten Einrichtung nutzen.', 'warn');
                    }
                });
                tdAction.appendChild(btnCreate);

                tr.append(tdKlasse, tdName, tdEmail, tdMs, tdAction);
                studentCouncilTbody.appendChild(tr);
            });
        }

        function setSingleGroupMatchStatus(el, payload, expectedNick) {
            if (!el) return;
            el.style.fontSize = '0.88em';
            el.style.lineHeight = '1.35';
            el.style.background = '';
            el.style.color = 'var(--muted)';
            el.title = '';
            if (!payload) {
                el.textContent = 'Noch nicht geprüft';
                updateSchoolWideGroupMatchActions();
                return;
            }
            if (payload.loading) {
                el.textContent = 'Prüfe …';
                return;
            }
            if (payload.found && payload.group) {
                const g = payload.group || {};
                const gid = normStr(g.id);
                const short = gid.length > 14 ? gid.slice(0, 12) + '…' : gid;
                const shown =
                    normStr(g.displayName) ||
                    normStr(g.mailNickname) ||
                    (expectedNick ? expectedNick : '') ||
                    'gematcht';
                el.innerHTML =
                    '<span style="color:#0d8050;font-weight:700;">✓</span> <code style="font-size:0.92em;">' +
                    escapeHtml(shown) +
                    '</code>';
                el.title =
                    (g.displayName ? String(g.displayName) : '') +
                    (g.mailNickname ? '\nAlias: ' + String(g.mailNickname) : '') +
                    (g.mail ? '\nMail: ' + String(g.mail) : '') +
                    (gid ? '\nObject-ID: ' + short : '');
                el.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
                el.style.color = 'var(--text)';
                updateSchoolWideGroupMatchActions();
                return;
            }
            if (payload.notFound) {
                el.innerHTML =
                    '<span style="color:#856404;font-weight:700;">✗</span> <span style="color:var(--muted)">nicht gefunden</span>';
                el.title = 'Keine passende Gruppe in Microsoft 365 gefunden';
                updateSchoolWideGroupMatchActions();
                return;
            }
            if (payload.error) {
                el.textContent = 'Fehler';
                el.title = String(payload.error || '');
                updateSchoolWideGroupMatchActions();
                return;
            }
            el.textContent = '–';
            updateSchoolWideGroupMatchActions();
        }

        function updateSchoolWideGroupMatchActions() {
            const sgaId = getMatchedGroupId('sga');
            const svId = getMatchedGroupId('studentCouncil');
            if (btnUnmatchSgaGroup) btnUnmatchSgaGroup.hidden = !sgaId;
            if (btnOpenSgaGroup) btnOpenSgaGroup.hidden = !sgaId;
            if (btnUnmatchStudentCouncilGroup) btnUnmatchStudentCouncilGroup.hidden = !svId;
            if (btnOpenStudentCouncilGroup) btnOpenStudentCouncilGroup.hidden = !svId;
        }

        function getMatchedGroupId(kind) {
            const api = window.ms365AppDataV2;
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const matched = setup && setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
            if (kind === 'sga') return normStr(matched.sgaGroupId || '');
            if (kind === 'studentCouncil') return normStr(matched.studentCouncilGroupId || '');
            return '';
        }

        function patchMatchedGroupId(kind, group) {
            const api = window.ms365AppDataV2;
            if (!api || typeof api.patchSetup !== 'function') return;
            const gid = normStr(group && group.id);
            const key = kind === 'sga' ? 'sgaGroupId' : 'studentCouncilGroupId';
            api.patchSetup({ matched: { [key]: gid || null } });
            updateSchoolWideGroupMatchActions();
        }

        function clearStoredSchoolWideGroupMatches() {
            patchMatchedGroupId('sga', null);
            patchMatchedGroupId('studentCouncil', null);
            setSingleGroupMatchStatus(sgaGroupMatchCell, null);
            setSingleGroupMatchStatus(studentCouncilGroupMatchCell, null);
            if (sgaGroupSearchResults) {
                sgaGroupSearchResults.style.display = 'none';
                sgaGroupSearchResults.replaceChildren();
            }
            if (studentCouncilGroupSearchResults) {
                studentCouncilGroupSearchResults.style.display = 'none';
                studentCouncilGroupSearchResults.replaceChildren();
            }
            renderStatusOverview();
        }

        function restoreStoredSchoolWideGroupMatchStatus() {
            const sgaId = getMatchedGroupId('sga');
            const svId = getMatchedGroupId('studentCouncil');
            if (sgaId) {
                setSingleGroupMatchStatus(
                    sgaGroupMatchCell,
                    {
                        found: true,
                        group: {
                            id: sgaId,
                            displayName: expectedSgaGroupDisplayName(),
                            mailNickname: expectedSgaGroupMailNickname()
                        }
                    },
                    expectedSgaGroupMailNickname()
                );
            } else {
                setSingleGroupMatchStatus(sgaGroupMatchCell, null);
            }
            if (svId) {
                setSingleGroupMatchStatus(
                    studentCouncilGroupMatchCell,
                    {
                        found: true,
                        group: {
                            id: svId,
                            displayName: expectedStudentCouncilGroupDisplayName(),
                            mailNickname: expectedStudentCouncilGroupMailNickname()
                        }
                    },
                    expectedStudentCouncilGroupMailNickname()
                );
            } else {
                setSingleGroupMatchStatus(studentCouncilGroupMatchCell, null);
            }
            updateSchoolWideGroupMatchActions();
        }

        async function resolveSchoolWideGroupByStoredOrExpected(kind, expectedNick, expectedDn) {
            const token = await graphApi().getGraphToken();
            const storedId = getMatchedGroupId(kind);
            if (storedId && typeof graphApi().fetchGroup === 'function') {
                try {
                    const g = await graphApi().fetchGroup(token, storedId);
                    if (g && g.id) return { found: true, group: g, by: 'stored' };
                } catch {
                    // gespeicherte ID ungültig → weiter mit Nick/Name-Suche
                }
            }
            const queries = [expectedNick, expectedDn].filter(Boolean);
            let hits = [];
            for (let i = 0; i < queries.length; i++) {
                hits = await graphApi().searchUnifiedGroups(token, queries[i]);
                if (hits && hits.length) break;
            }
            const nickLc = normStr(expectedNick).toLowerCase();
            const matchByNick =
                Array.isArray(hits) &&
                hits.find(function (g) {
                    const mn = normStr(g && g.mailNickname).toLowerCase();
                    if (mn && mn === nickLc) return true;
                    const mail = normStr(g && g.mail).toLowerCase();
                    return mail && mail.startsWith(nickLc + '@');
                });
            if (matchByNick) return { found: true, group: matchByNick, by: 'nick' };
            return { found: false, notFound: true, hits: hits || [] };
        }

        function renderSchoolWideGroupSearchResults(host, kind, list) {
            if (!host) return;
            host.replaceChildren();
            if (!list || !list.length) {
                host.style.display = 'none';
                return;
            }
            host.style.display = 'block';
            const box = document.createElement('div');
            box.style.border = '1px solid var(--border)';
            box.style.borderRadius = '10px';
            box.style.overflow = 'hidden';
            list.forEach(function (g, idx) {
                const row = document.createElement('div');
                row.style.display = 'flex';
                row.style.flexWrap = 'wrap';
                row.style.gap = '10px';
                row.style.alignItems = 'center';
                row.style.justifyContent = 'space-between';
                row.style.padding = '10px 12px';
                if (idx > 0) row.style.borderTop = '1px solid var(--border)';
                const dn = normStr(g && g.displayName) || '(ohne Namen)';
                const mail = normStr(g && g.mail) || '–';
                const nick = normStr(g && g.mailNickname) || '–';
                const left = document.createElement('div');
                left.style.minWidth = '0';
                left.style.flex = '1 1 220px';
                left.innerHTML =
                    '<div style="font-weight:700;line-height:1.25;">' +
                    escapeHtml(dn) +
                    '</div>' +
                    '<div class="muted" style="margin-top:2px;font-size:0.92em;">Alias: <code>' +
                    escapeHtml(nick) +
                    '</code> · ' +
                    escapeHtml(mail) +
                    '</div>';
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn btn-primary';
                btn.innerHTML = '<i class="bi bi-link-45deg" aria-hidden="true"></i> Zuordnen';
                btn.addEventListener('click', function () {
                    applyManualSchoolWideGroupMatch(kind, g);
                });
                row.appendChild(left);
                row.appendChild(btn);
                box.appendChild(row);
            });
            host.appendChild(box);
        }

        function applyManualSchoolWideGroupMatch(kind, group) {
            if (!group || !group.id) return;
            const expectedNick =
                kind === 'sga' ? expectedSgaGroupMailNickname() : expectedStudentCouncilGroupMailNickname();
            const cell = kind === 'sga' ? sgaGroupMatchCell : studentCouncilGroupMatchCell;
            const resultsHost = kind === 'sga' ? sgaGroupSearchResults : studentCouncilGroupSearchResults;
            patchMatchedGroupId(kind, group);
            setSingleGroupMatchStatus(cell, { found: true, group: group }, expectedNick);
            if (resultsHost) {
                resultsHost.style.display = 'none';
                resultsHost.replaceChildren();
            }
            setSummary(
                (kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': bestehende Microsoft-365-Gruppe zugeordnet.',
                'ok'
            );
            renderStatusOverview();
        }

        async function searchSchoolWideGroups(kind) {
            const inp = kind === 'sga' ? inpSgaGroupSearch : inpStudentCouncilGroupSearch;
            const host = kind === 'sga' ? sgaGroupSearchResults : studentCouncilGroupSearchResults;
            const q = inp && inp.value ? String(inp.value).trim() : '';
            if (!q) {
                setSummary('Bitte einen Suchbegriff eingeben (Name, Mail oder Alias).', 'warn');
                return;
            }
            setSummary((kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': Suche Gruppen …', 'warn');
            const token = await graphApi().getGraphToken();
            const list = await graphApi().searchUnifiedGroups(token, q);
            renderSchoolWideGroupSearchResults(host, kind, list);
            if (!list || !list.length) {
                setSummary('Keine passenden Microsoft-365-Gruppen gefunden.', 'warn');
            } else {
                setSummary(list.length + ' Gruppe(n) gefunden – bitte zuordnen.', 'ok');
            }
        }

        function unmatchSchoolWideGroup(kind) {
            const cell = kind === 'sga' ? sgaGroupMatchCell : studentCouncilGroupMatchCell;
            const host = kind === 'sga' ? sgaGroupSearchResults : studentCouncilGroupSearchResults;
            patchMatchedGroupId(kind, null);
            setSingleGroupMatchStatus(cell, null);
            if (host) {
                host.style.display = 'none';
                host.replaceChildren();
            }
            setSummary((kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': Match gelöst.', 'ok');
            renderStatusOverview();
        }

        function openSchoolWideGroupInEntra(kind) {
            const gid = getMatchedGroupId(kind);
            if (!gid) {
                setSummary('Keine gematchte Gruppe zum Öffnen.', 'warn');
                return;
            }
            window.open(ENTRA_GROUP_DETAILS + encodeURIComponent(gid), '_blank', 'noopener');
        }

        function schoolBaseNick() {
            const nm = schoolNameInput ? normStr(schoolNameInput.value || '') : '';
            return nm.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().slice(0, 24);
        }

        function expectedSgaGroupMailNickname() {
            const base = schoolBaseNick();
            if (!base) return '';
            return graphApi().sanitizeUnifiedGroupMailNickname('sga' + base);
        }

        function expectedSgaGroupDisplayName() {
            const nm = schoolNameInput ? normStr(schoolNameInput.value || '') : '';
            return nm ? 'SGA ' + nm : '';
        }

        function expectedStudentCouncilGroupMailNickname() {
            const base = schoolBaseNick();
            if (!base) return '';
            const yLbl = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            const digits = String(yLbl || '').replace(/\D/g, '').slice(0, 6);
            return graphApi().sanitizeUnifiedGroupMailNickname('sv' + digits + base);
        }

        function expectedStudentCouncilGroupDisplayName() {
            const nm = schoolNameInput ? normStr(schoolNameInput.value || '') : '';
            const yLbl = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            if (!nm) return '';
            return 'Schülervertretung ' + nm + ' ' + yLbl;
        }

        async function verifySgaGroupExistence() {
            if (!btnVerifySgaGraph) {
                // btnVerifySgaGraph existiert als „Guard“ für die Seite – wenn nicht, ist die UI nicht aktiv.
                throw new Error('UI nicht bereit.');
            }
            const expectedNick = expectedSgaGroupMailNickname();
            const expectedDn = expectedSgaGroupDisplayName();
            if (!expectedNick || !expectedDn) {
                // Ohne Schulname trotzdem gespeicherte Zuordnung behalten/anzeigen
                const storedOnly = getMatchedGroupId('sga');
                if (storedOnly) {
                    setSingleGroupMatchStatus(
                        sgaGroupMatchCell,
                        { found: true, group: { id: storedOnly } },
                        ''
                    );
                    return { found: true, group: { id: storedOnly } };
                }
                patchMatchedGroupId('sga', null);
                setSingleGroupMatchStatus(sgaGroupMatchCell, { notFound: true }, expectedNick || '');
                setSummary('SGA: Bitte zuerst „Schulname“ setzen.', 'warn');
                return { found: false, notFound: true };
            }
            setSingleGroupMatchStatus(sgaGroupMatchCell, { loading: true }, expectedNick);
            const resolved = await resolveSchoolWideGroupByStoredOrExpected('sga', expectedNick, expectedDn);
            if (resolved && resolved.found && resolved.group) {
                patchMatchedGroupId('sga', resolved.group);
                setSingleGroupMatchStatus(sgaGroupMatchCell, { found: true, group: resolved.group }, expectedNick);
                return { found: true, group: resolved.group };
            }
            patchMatchedGroupId('sga', null);
            setSingleGroupMatchStatus(sgaGroupMatchCell, { notFound: true }, expectedNick);
            return { found: false, notFound: true };
        }

        async function createSgaGroupExistence() {
            const expectedNick = expectedSgaGroupMailNickname();
            const expectedDn = expectedSgaGroupDisplayName();
            if (!expectedNick || !expectedDn) {
                setSummary('SGA: Bitte zuerst „Schulname“ setzen.', 'warn');
                patchMatchedGroupId('sga', null);
                setSingleGroupMatchStatus(sgaGroupMatchCell, { notFound: true }, expectedNick || '');
                return null;
            }
            const mode = selSgaMode ? normStr(selSgaMode.value || 'group').toLowerCase() : 'group';

            // Vermeidet Dubletten: Erst prüfen, dann ggf. anlegen.
            const existing = await verifySgaGroupExistence().catch(function () {
                return { found: false };
            });
            if (existing && existing.found && existing.group && existing.group.id) return existing.group;

            setSingleGroupMatchStatus(sgaGroupMatchCell, { loading: true }, expectedNick);
            const token = await graphApi().getGraphToken();
            const created = await graphApi().createUnifiedGroup(token, expectedDn, expectedNick, 'MS365-Schul-Tools – SGA');
            if (mode === 'group') {
                try {
                    await graphApi().provisionTeamForGroup(token, created && created.id ? created.id : '');
                } catch {
                    // optional: Team-Provision kann warten/fehlschlagen
                }
            }
            const match = created && created.id ? { id: created.id, displayName: created.displayName || expectedDn, mailNickname: created.mailNickname || expectedNick } : null;
            patchMatchedGroupId('sga', match || created);
            setSingleGroupMatchStatus(sgaGroupMatchCell, { found: true, group: match || created }, expectedNick);
            return created;
        }

        async function verifyStudentCouncilGroupExistence() {
            const expectedNick = expectedStudentCouncilGroupMailNickname();
            const expectedDn = expectedStudentCouncilGroupDisplayName();
            if (!expectedNick || !expectedDn) {
                const storedOnly = getMatchedGroupId('studentCouncil');
                if (storedOnly) {
                    setSingleGroupMatchStatus(
                        studentCouncilGroupMatchCell,
                        { found: true, group: { id: storedOnly } },
                        ''
                    );
                    return { found: true, group: { id: storedOnly } };
                }
                patchMatchedGroupId('studentCouncil', null);
                setSingleGroupMatchStatus(studentCouncilGroupMatchCell, { notFound: true }, expectedNick || '');
                setSummary('Schülervertretung: Bitte zuerst „Schulname“ setzen.', 'warn');
                return { found: false, notFound: true };
            }
            setSingleGroupMatchStatus(studentCouncilGroupMatchCell, { loading: true }, expectedNick);
            const resolved = await resolveSchoolWideGroupByStoredOrExpected(
                'studentCouncil',
                expectedNick,
                expectedDn
            );
            if (resolved && resolved.found && resolved.group) {
                patchMatchedGroupId('studentCouncil', resolved.group);
                setSingleGroupMatchStatus(
                    studentCouncilGroupMatchCell,
                    { found: true, group: resolved.group },
                    expectedNick
                );
                return { found: true, group: resolved.group };
            }
            patchMatchedGroupId('studentCouncil', null);
            setSingleGroupMatchStatus(studentCouncilGroupMatchCell, { notFound: true }, expectedNick);
            return { found: false, notFound: true };
        }

        async function createStudentCouncilGroupExistence() {
            const expectedNick = expectedStudentCouncilGroupMailNickname();
            const expectedDn = expectedStudentCouncilGroupDisplayName();
            if (!expectedNick || !expectedDn) {
                setSummary('Schülervertretung: Bitte zuerst „Schulname“ setzen.', 'warn');
                patchMatchedGroupId('studentCouncil', null);
                setSingleGroupMatchStatus(studentCouncilGroupMatchCell, { notFound: true }, expectedNick || '');
                return null;
            }

            const existing = await verifyStudentCouncilGroupExistence().catch(function () {
                return { found: false };
            });
            if (existing && existing.found && existing.group && existing.group.id) return existing.group;

            setSingleGroupMatchStatus(studentCouncilGroupMatchCell, { loading: true }, expectedNick);
            const token = await graphApi().getGraphToken();
            const created = await graphApi().createUnifiedGroup(
                token,
                expectedDn,
                expectedNick,
                'MS365-Schul-Tools – Schülervertretung'
            );
            // Schülervertretung: Team-Provision überspringen (kein explizites Team-Zielbild in den aktuellen Anforderungen)
            const match =
                created && created.id
                    ? { id: created.id, displayName: created.displayName || expectedDn, mailNickname: created.mailNickname || expectedNick }
                    : null;
            patchMatchedGroupId('studentCouncil', match || created);
            setSingleGroupMatchStatus(studentCouncilGroupMatchCell, { found: true, group: match || created }, expectedNick);
            return created;
        }

        let lastLifecyclePreview = null;
        let pendingSisImport = null;

        function sisApi() {
            return window.ms365SchoolSisImport || null;
        }

        function hideSisDiff() {
            pendingSisImport = null;
            const box = document.getElementById('tenantSisDiff');
            if (box) box.hidden = true;
        }

        function rememberSisHistory(diff, source, mode) {
            try {
                const api = window.ms365AppDataV2;
                if (!api || typeof api.getSetup !== 'function' || typeof api.patchSetup !== 'function') return;
                const setup = api.getSetup() || {};
                const hist = Array.isArray(setup.sisImportHistory) ? setup.sisImportHistory.slice() : [];
                const c = (diff && diff.counts) || {};
                hist.push({
                    at: new Date().toISOString(),
                    source: source || '',
                    mode: mode,
                    added: c.added || 0,
                    updated: c.updated || 0,
                    removed: c.removed || 0,
                    conflicts: c.conflicts || 0
                });
                api.patchSetup({ sisImportHistory: hist.slice(-20) });
            } catch {
                /* ignore */
            }
            if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                window.ms365ActionLog.append({
                    tool: 'sis-import',
                    action: mode === 'replace' ? 'replace' : 'merge',
                    summary: (sisApi() && sisApi().summarizeSisDiff ? sisApi().summarizeSisDiff(diff) : '') + (source ? ' · ' + source : '')
                });
            }
        }

        function applyPendingSis(mode) {
            const pending = pendingSisImport;
            const sis = sisApi();
            if (!pending || !sis || typeof sis.applySisImport !== 'function') return;
            const next = sis.applySisImport(pending.existing, pending.records, { mode: mode });
            if (taStudents) taStudents.value = sis.recordsToSemicolonLines(next);
            renderStudentsTableFromTextarea();
            const ySt = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            setSummary(
                'Schülerimport übernommen (' +
                    (mode === 'replace' ? 'ersetzt' : 'zusammengeführt') +
                    '): ' +
                    next.length +
                    ' Zeilen — Schuljahr ' +
                    ySt,
                'ok'
            );
            rememberSisHistory(pending.diff, pending.source, mode);
            renderStudentLifecyclePanel(pending.existing, getStudentsFromTextarea());
            hideSisDiff();
        }

        function showSisDiffPreview(existing, result) {
            const sis = sisApi();
            const box = document.getElementById('tenantSisDiff');
            const meta = document.getElementById('tenantSisDiffMeta');
            const ul = document.getElementById('tenantSisDiffList');
            if (!sis || typeof sis.diffSisImport !== 'function' || !box) {
                if (taStudents) taStudents.value = (result && result.lines) || '';
                renderStudentsTableFromTextarea();
                return;
            }
            const records = result && Array.isArray(result.records) ? result.records : [];
            const diff = sis.diffSisImport(existing || [], records);
            pendingSisImport = {
                existing: existing || [],
                records: records,
                source: result && result.source ? result.source : '',
                diff: diff
            };
            if (meta) {
                const extra = [];
                if (result && result.meta && result.meta.emailsGenerated) {
                    extra.push(result.meta.emailsGenerated + ' Mails aus Namen vorgeschlagen');
                }
                if (result && result.meta && result.meta.unmatchedGuardians) {
                    extra.push(result.meta.unmatchedGuardians + ' Eltern ohne Schüler-Zuordnung');
                }
                meta.textContent =
                    (sis.summarizeSisDiff ? sis.summarizeSisDiff(diff) : '') +
                    (result && result.source ? ' · Quelle ' + result.source : '') +
                    (extra.length ? ' · ' + extra.join(' · ') : '') +
                    '. Zusammenführen behält lokale Zeilen, die in der Datei fehlen (geeignet für Aktualisierungen im Jahr).';
            }
            if (ul) {
                ul.innerHTML = '';
                (diff.conflicts || []).slice(0, 8).forEach(function (c) {
                    const li = document.createElement('li');
                    li.textContent = 'Konflikt: ' + (c.summary || c.type);
                    ul.appendChild(li);
                });
                (diff.updated || []).slice(0, 6).forEach(function (u) {
                    const li = document.createElement('li');
                    const prev = u.previous || {};
                    const rec = u.incoming || {};
                    li.textContent =
                        'Änderung: ' +
                        (prev.name || prev.email || '') +
                        (u.klasseChanged ? ' (' + (prev.klasse || '–') + ' → ' + (rec.klasse || '–') + ')' : '');
                    ul.appendChild(li);
                });
                if (!ul.childNodes.length) {
                    const li = document.createElement('li');
                    li.textContent = (diff.counts && diff.counts.added ? diff.counts.added + ' neue Zeilen. ' : '') + 'Keine Konflikte erkannt.';
                    ul.appendChild(li);
                }
            }
            box.hidden = false;
        }

        function lifecycleApi() {
            return window.ms365StudentClassLifecycle || null;
        }

        function getLifecycleContext() {
            const empty = { classTeams: [], schuelerGroupId: '' };
            try {
                const api = window.ms365AppDataV2;
                const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
                const classTeams = c && c.core && Array.isArray(c.core.classTeams) ? c.core.classTeams : [];
                const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : c && c.setup;
                const matched = setup && setup.matched ? setup.matched : {};
                return {
                    classTeams: classTeams,
                    schuelerGroupId: matched.schuelerGroupId ? String(matched.schuelerGroupId) : ''
                };
            } catch {
                return empty;
            }
        }

        function renderStudentLifecyclePanel(prev, next) {
            const host = document.getElementById('tenantStudentLifecycle');
            const listEl = document.getElementById('tenantStudentLifecycleList');
            const metaEl = document.getElementById('tenantStudentLifecycleMeta');
            const lc = lifecycleApi();
            if (!host || !lc) return;
            const diff = lc.diffStudents(prev, next);
            const ctx = getLifecycleContext();
            const preview = lc.previewMemberships(diff, ctx.classTeams, ctx.schuelerGroupId);
            lastLifecyclePreview = preview;
            if (!lc.hasMembershipWork(preview)) {
                host.hidden = true;
                if (listEl) listEl.replaceChildren();
                return;
            }
            const sum = lc.summarizePreview(preview);
            host.hidden = false;
            if (metaEl) {
                metaEl.textContent =
                    sum.join +
                    ' Aufnahme(n), ' +
                    sum.leave +
                    ' Entfernung(en) in ' +
                    sum.groupCount +
                    ' Gruppe(n). Nur E-Mails mit zugeordneter Microsoft-365-Gruppe erscheinen hier.';
            }
            if (listEl) {
                listEl.replaceChildren();
                preview.groups.forEach(function (g) {
                    const li = document.createElement('li');
                    const joinN = g.join.length;
                    const leaveN = g.leave.length;
                    li.textContent =
                        (g.label || g.groupId) +
                        ': +' +
                        joinN +
                        (joinN ? ' (' + g.join.slice(0, 4).join(', ') + (joinN > 4 ? ' …' : '') + ')' : '') +
                        ', −' +
                        leaveN +
                        (leaveN ? ' (' + g.leave.slice(0, 4).join(', ') + (leaveN > 4 ? ' …' : '') + ')' : '');
                    listEl.appendChild(li);
                });
            }
        }

        async function applyStudentLifecyclePreview() {
            const lc = lifecycleApi();
            const gug = window.ms365GraphUnifiedGroups;
            if (!lc || !gug || !lastLifecyclePreview || !lc.hasMembershipWork(lastLifecyclePreview)) {
                setSummary('Keine Mitgliedschaftsänderungen zum Anwenden.', 'warn');
                return;
            }
            const btn = document.getElementById('tenantStudentLifecycleApply');
            if (btn) btn.disabled = true;
            try {
                const token = await gug.getGraphToken();
                let ok = 0;
                let fail = 0;
                for (let i = 0; i < lastLifecyclePreview.groups.length; i++) {
                    const g = lastLifecyclePreview.groups[i];
                    if (g.join && g.join.length) {
                        const r = await gug.syncEmailsToGroup(token, g.groupId, g.join, g.label, function () {});
                        ok += r.ok || 0;
                        fail += r.fail || 0;
                    }
                    if (g.leave && g.leave.length && typeof gug.removeEmailsFromGroup === 'function') {
                        const r = await gug.removeEmailsFromGroup(token, g.groupId, g.leave, g.label, function () {});
                        ok += r.ok || 0;
                        fail += r.fail || 0;
                    }
                }
                setSummary(
                    'Mitgliedschaften in Microsoft 365 angewendet: ' + ok + ' OK' + (fail ? ', ' + fail + ' Fehler' : '') + '.',
                    fail ? 'warn' : 'ok'
                );
                const host = document.getElementById('tenantStudentLifecycle');
                if (host) host.hidden = true;
            } catch (e) {
                setSummary('Mitgliedschaften: ' + (e.message || e), 'warn');
            } finally {
                if (btn) btn.disabled = false;
            }
        }

        function setStudentsTextareaFromRows(rows) {
            if (!taStudents) return;
            taStudents.value = studentsToLines(rows);
        }

        function getStudentsDirectoryMatchMap() {
            try {
                const api = window.ms365AppDataV2;
                const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
                return setup && setup.directoryMatchByEmail ? setup.directoryMatchByEmail : {};
            } catch {
                return {};
            }
        }

        /** Update one student field without rebuilding all ~N table rows. */
        function commitStudentCellEdit(idx, field, next, meta, td) {
            const all = getStudentsFromTextarea();
            if (!all[idx]) {
                renderStudentsTableFromTextarea();
                return;
            }
            const prev = all[idx][field];
            let value = prev;
            if (!(meta && meta.cancelled)) {
                if (field === 'email') value = normStr(next).toLowerCase();
                else value = normStr(next);
            }
            all[idx][field] = value;
            setStudentsTextareaFromRows(all);

            const filterActive =
                (studentsClassFilter && studentsClassFilter.value) ||
                (studentsTextFilter && String(studentsTextFilter.value || '').trim());
            if (field === 'klasse' || filterActive) {
                renderStudentsTableFromTextarea();
                scheduleAutoSave();
                return;
            }

            td.textContent = value || '';
            td.title = 'Doppelklick zum Bearbeiten';

            if (field === 'email') {
                const tr = td.parentElement;
                if (tr) {
                    const oldMs = tr.children[5];
                    const newMs = buildDirectoryMatchCell(tr, value, getStudentsDirectoryMatchMap());
                    if (oldMs) tr.replaceChild(newMs, oldMs);
                    else tr.appendChild(newMs);
                    tr.style.background = '';
                    const match = value ? getDirectoryMatchByEmail(value) : null;
                    if (match && match.graphUserId) {
                        tr.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
                    }
                }
            }
            scheduleAutoSave();
        }

        function updateStudentsSortIndicators() {
            updateTableSortIndicators(studentsTable, studentsSortState);
        }

        function applySubjectsSort(key) {
            if (!key || (key !== 'code' && key !== 'name')) return;
            toggleTenantSortState(subjectsSortState, key);
            setSubjectsTextareaFromRows(sortTenantRows(getSubjectsFromTextarea(), key, subjectsSortState.dir));
            renderSubjectsTableFromTextarea();
            scheduleAutoSave();
        }

        function applyArgesSort(key) {
            if (!key || (key !== 'code' && key !== 'name' && key !== 'subjects')) return;
            toggleTenantSortState(argesSortState, key);
            setArgesTextareaFromRows(sortTenantRows(getArgesFromTextarea(), key, argesSortState.dir));
            renderArgesTableFromTextarea();
            scheduleAutoSave();
        }

        function applyClassesSort(key) {
            const allowed = ['code', 'year', 'name', 'headName', 'headEmail'];
            if (!key || allowed.indexOf(key) < 0) return;
            toggleTenantSortState(classesSortState, key);
            setClassesTextareaFromRows(sortTenantRows(getClassesFromTextarea(), key, classesSortState.dir));
            renderClassesTableFromTextarea();
            scheduleAutoSave();
        }

        function applySgaSort(key) {
            if (!key || (key !== 'scope' && key !== 'name' && key !== 'email')) return;
            toggleTenantSortState(sgaSortState, key);
            setSgaTextareaFromRows(sortTenantRows(getSgaFromTextarea(), key, sgaSortState.dir));
            renderSgaTableFromTextarea();
            scheduleAutoSave();
        }

        function applyStudentCouncilSort(key) {
            if (!key || (key !== 'klasse' && key !== 'name' && key !== 'email')) return;
            toggleTenantSortState(studentCouncilSortState, key);
            setStudentCouncilTextareaFromRows(
                sortTenantRows(getStudentCouncilFromTextarea(), key, studentCouncilSortState.dir)
            );
            renderStudentCouncilTableFromTextarea();
            scheduleAutoSave();
        }

        function applyAdminUnifiedSort(key) {
            const allowed = ['name', 'code', 'personName', 'email'];
            if (!key || allowed.indexOf(key) < 0) return;
            toggleTenantSortState(adminUnifiedSortState, key);
            const sorted = sortTenantRows(groupsToDisplayRows(getAdministrationGroups()), key, adminUnifiedSortState.dir);
            setAdministrationGroups(displayRowsToGroups(sorted));
            renderAdminUnifiedTableFromBundle();
            scheduleAutoSave();
        }

        function syncStudentsClassFilterOptions(rows) {
            if (!studentsClassFilter) return;
            const prev = String(studentsClassFilter.value || '');
            const codes = new Set();
            let hasEmpty = false;
            (rows || []).forEach((row) => {
                const k = normStr(row && row.klasse);
                if (!k) hasEmpty = true;
                else codes.add(k);
            });
            const sorted = Array.from(codes).sort((a, b) => a.localeCompare(b, 'de', { sensitivity: 'base' }));
            const signature = (hasEmpty ? '1' : '0') + '|' + sorted.join('\n');
            if (studentsClassFilter.dataset.sig !== signature) {
                studentsClassFilter.dataset.sig = signature;
                studentsClassFilter.replaceChildren();
                const optAll = document.createElement('option');
                optAll.value = '';
                optAll.textContent = 'Alle Klassen';
                studentsClassFilter.appendChild(optAll);
                if (hasEmpty) {
                    const optEmpty = document.createElement('option');
                    optEmpty.value = '__empty__';
                    optEmpty.textContent = 'Ohne Klasse';
                    studentsClassFilter.appendChild(optEmpty);
                }
                sorted.forEach((code) => {
                    const opt = document.createElement('option');
                    opt.value = code;
                    opt.textContent = code;
                    studentsClassFilter.appendChild(opt);
                });
            }
            let next = '';
            if (prev === '__empty__' && hasEmpty) {
                next = '__empty__';
            } else if (prev) {
                const match = sorted.find((c) => c.toUpperCase() === prev.toUpperCase());
                if (match) next = match;
            }
            if (studentsClassFilter.value !== next) {
                studentsClassFilter.value = next;
            }
        }

        function studentRowMatchesFilters(row, classFilter, textQuery) {
            const klasse = normStr(row && row.klasse);
            if (classFilter === '__empty__') {
                if (klasse) return false;
            } else if (classFilter) {
                if (klasse.toUpperCase() !== classFilter.toUpperCase()) return false;
            }
            const tokens = String(textQuery || '')
                .trim()
                .toLowerCase()
                .split(/\s+/)
                .filter(Boolean);
            if (!tokens.length) return true;
            const parts = [klasse, row && row.name, row && row.email];
            const pairs = Array.isArray(row && row.parentPairs) ? row.parentPairs : [];
            pairs.forEach((p) => {
                if (!p) return;
                parts.push(p.name, p.email);
            });
            const hay = parts
                .map((v) => String(v || '').toLowerCase())
                .join('\n');
            return tokens.every((t) => hay.includes(t));
        }

        function updateStudentsFilterMeta(total, visible) {
            if (!studentsFilterMeta) return;
            const classActive = studentsClassFilter && studentsClassFilter.value;
            const textActive = studentsTextFilter && String(studentsTextFilter.value || '').trim();
            if (!total || (!classActive && !textActive)) {
                studentsFilterMeta.hidden = true;
                studentsFilterMeta.textContent = '';
                return;
            }
            studentsFilterMeta.hidden = false;
            studentsFilterMeta.textContent =
                visible === total
                    ? total + ' Einträge angezeigt'
                    : visible + ' von ' + total + ' Einträgen angezeigt';
        }

        function applyStudentsSort(key) {
            if (!key || (key !== 'klasse' && key !== 'name' && key !== 'email')) return;
            if (studentsSortState.key === key) {
                studentsSortState.dir = studentsSortState.dir === 1 ? -1 : 1;
            } else {
                studentsSortState.key = key;
                studentsSortState.dir = 1;
            }
            const api = window.ms365StudentListImport;
            const all = getStudentsFromTextarea();
            const sorted =
                api && typeof api.sortRows === 'function'
                    ? api.sortRows(all, studentsSortState.key, studentsSortState.dir)
                    : all.slice().sort((a, b) => {
                          const av = String((a && a[key]) || '');
                          const bv = String((b && b[key]) || '');
                          return av.localeCompare(bv, 'de', { sensitivity: 'base' }) * studentsSortState.dir;
                      });
            setStudentsTextareaFromRows(sorted);
            renderStudentsTableFromTextarea();
            scheduleAutoSave();
        }

        function renderStudentsTableFromTextarea() {
            if (!studentsTbody) return;
            const rows = getStudentsFromTextarea();
            updateStudentsSortIndicators();
            syncStudentsClassFilterOptions(rows);
            const classFilter = studentsClassFilter ? String(studentsClassFilter.value || '') : '';
            const textQuery = studentsTextFilter ? String(studentsTextFilter.value || '') : '';
            const visible = [];
            rows.forEach((row, idx) => {
                if (studentRowMatchesFilters(row, classFilter, textQuery)) {
                    visible.push({ row, idx });
                }
            });
            updateStudentsFilterMeta(rows.length, visible.length);

            if (!rows.length) {
                studentsTbody.replaceChildren();
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 7;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen, aus Microsoft 365 einlesen oder „+ Zeile“.';
                tr.appendChild(td);
                studentsTbody.appendChild(tr);
                return;
            }

            if (!visible.length) {
                studentsTbody.replaceChildren();
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 7;
                td.style.color = 'var(--muted)';
                td.textContent = 'Keine Treffer für Klasse/Suche – Filter anpassen.';
                tr.appendChild(td);
                studentsTbody.appendChild(tr);
                return;
            }

            const dirMatchMap = getStudentsDirectoryMatchMap();

            function buildParentCell(pair, extraCount) {
                const td = document.createElement('td');
                td.className = 'students-parent-cell';
                const p = pair || {};
                const pName = normStr(p.name || '');
                const pMail = normStr(p.email || '').toLowerCase();
                if (!pName && !pMail) {
                    td.textContent = '–';
                    td.style.color = 'var(--muted)';
                    td.title = 'Kein Erziehungsberechtigter';
                    return td;
                }
                td.textContent = pName || pMail;
                if (extraCount > 0) {
                    const more = document.createElement('span');
                    more.className = 'students-parent-more';
                    more.textContent = ' +' + extraCount;
                    td.appendChild(more);
                }
                const tipParts = [];
                if (pName && pMail) tipParts.push(pName);
                if (pMail) tipParts.push(pMail);
                else if (pName) tipParts.push(pName);
                if (extraCount > 0) tipParts.push('(+ ' + extraCount + ' weitere)');
                td.title = tipParts.join(' · ') || 'Erziehungsberechtigte:r';
                return td;
            }

            const frag = document.createDocumentFragment();
            visible.forEach(({ row, idx }) => {
                const tr = document.createElement('tr');
                tr.dataset.studentIdx = String(idx);

                const tdClass = document.createElement('td');
                tdClass.innerHTML = '<code>' + escapeHtml(row.klasse || '') + '</code>';
                tdClass.title = 'Doppelklick zum Bearbeiten';
                tdClass.addEventListener('dblclick', () => {
                    startCellEdit(tdClass, tdClass.textContent, (next, meta) => {
                        commitStudentCellEdit(idx, 'klasse', next, meta, tdClass);
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, tdName.textContent, (next, meta) => {
                        commitStudentCellEdit(idx, 'name', next, meta, tdName);
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.email || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, tdEmail.textContent, (next, meta) => {
                        commitStudentCellEdit(idx, 'email', next, meta, tdEmail);
                    });
                });

                const pairs = Array.isArray(row.parentPairs) ? row.parentPairs : [];
                const tdParent1 = buildParentCell(pairs[0], 0);
                const tdParent2 = buildParentCell(pairs[1], Math.max(0, pairs.length - 2));

                const tdMs = buildDirectoryMatchCell(tr, row.email, dirMatchMap);

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                tdAction.style.whiteSpace = 'nowrap';
                tdAction.style.display = 'flex';
                tdAction.style.gap = '6px';
                tdAction.style.alignItems = 'center';
                const emKey = normStr(row.email || '').toLowerCase();
                const dir = emKey ? dirMatchMap[emKey] || null : null;
                if (row.email && row.name) {
                    const btnCheck = document.createElement('button');
                    btnCheck.type = 'button';
                    btnCheck.className = 'mini-btn';
                    btnCheck.style.background = 'var(--brand1)';
                    btnCheck.title = 'Diese E‑Mail in Microsoft Entra prüfen';
                    btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCheck.disabled = true;
                    } else {
                        btnCheck.addEventListener('click', async () => {
                            btnCheck.disabled = true;
                            try {
                                await verifyAdminDirectoryEmail(row.email);
                            } finally {
                                renderStudentsTableFromTextarea();
                            }
                        });
                    }
                    tdAction.appendChild(btnCheck);

                    const btnCreate = document.createElement('button');
                    btnCreate.type = 'button';
                    btnCreate.className = 'mini-btn';
                    btnCreate.style.background = '#11cdef';
                    btnCreate.title = 'Neuen Benutzer in Microsoft Entra ID anlegen (User.ReadWrite.All)';
                    btnCreate.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
                    if (dir && dir.graphUserId) {
                        btnCreate.disabled = true;
                    } else {
                        btnCreate.addEventListener('click', async () => {
                            btnCreate.disabled = true;
                            try {
                                await createAdminEntraUserInteractive(row.email, row.name);
                            } finally {
                                renderStudentsTableFromTextarea();
                            }
                        });
                    }
                    tdAction.appendChild(btnCreate);
                }
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
                    scheduleAutoSave();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdClass, tdName, tdEmail, tdParent1, tdParent2, tdMs, tdAction);
                frag.appendChild(tr);
            });
            studentsTbody.replaceChildren(frag);
        }

        function classesToLines(rows) {
            return classesToLinesPure(rows);
        }

        function getClassesFromTextarea() {
            return typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
        }

        function setClassesTextareaFromRows(rows) {
            if (!taClasses) return;
            taClasses.value = classesToLines(rows);
        }

        function renderClassesTableFromTextarea() {
            if (!classesTbody) return;
            updateTableSortIndicators(classesTable, classesSortState);
            const rows = getClassesFromTextarea();
            classesTbody.replaceChildren();

            if (!rows.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 7;
                td.style.color = 'var(--muted)';
                td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
                tr.appendChild(td);
                classesTbody.appendChild(tr);
                return;
            }

            rows.forEach((row, idx) => {
                const tr = document.createElement('tr');

                const tdCode = document.createElement('td');
                tdCode.innerHTML = `<code>${row.code || ''}</code>`;
                tdCode.title = 'Doppelklick zum Bearbeiten';
                tdCode.addEventListener('dblclick', () => {
                    startCellEdit(tdCode, row.code, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].code;
                        all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdYear = document.createElement('td');
                tdYear.textContent = row.year || '';
                tdYear.title = 'Doppelklick zum Bearbeiten';
                tdYear.addEventListener('dblclick', () => {
                    startCellEdit(tdYear, row.year, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].year || '';
                        const n = normStr(next);
                        all[idx].year = meta && meta.cancelled ? prev : /^\d{4}$/.test(n) ? n : '';
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdName = document.createElement('td');
                tdName.textContent = row.name || '';
                tdName.title = 'Doppelklick zum Bearbeiten';
                tdName.addEventListener('dblclick', () => {
                    startCellEdit(tdName, row.name, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].name;
                        all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdHead = document.createElement('td');
                tdHead.textContent = row.headName || '';
                tdHead.title = 'Doppelklick zum Bearbeiten';
                tdHead.addEventListener('dblclick', () => {
                    startCellEdit(tdHead, row.headName, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].headName;
                        all[idx].headName = meta && meta.cancelled ? prev : normStr(next);
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdEmail = document.createElement('td');
                tdEmail.textContent = row.headEmail || '';
                tdEmail.title = 'Doppelklick zum Bearbeiten';
                tdEmail.addEventListener('dblclick', () => {
                    startCellEdit(tdEmail, row.headEmail, (next, meta) => {
                        const all = getClassesFromTextarea();
                        if (!all[idx]) return renderClassesTableFromTextarea();
                        const prev = all[idx].headEmail;
                        all[idx].headEmail = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                        setClassesTextareaFromRows(all);
                        renderClassesTableFromTextarea();
                    });
                });

                const tdMs = buildClassGroupMatchCell(tr, row);

                const tdAction = document.createElement('td');
                tdAction.className = 'action-cell';
                tdAction.style.whiteSpace = 'nowrap';
                tdAction.style.display = 'flex';
                tdAction.style.gap = '6px';
                tdAction.style.alignItems = 'center';
                const btnCheck = document.createElement('button');
                btnCheck.type = 'button';
                btnCheck.className = 'mini-btn';
                btnCheck.style.background = 'var(--brand1)';
                btnCheck.title = 'Klassengruppe in Microsoft 365 prüfen';
                btnCheck.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
                btnCheck.addEventListener('click', async () => {
                    btnCheck.disabled = true;
                    try {
                        const res = await verifyClassGroupForRow(row);
                        setSummary(
                            (row.code || row.name || 'Klasse') + ': ' + (res && res.found ? 'Klassengruppe gefunden' : 'keine Klassengruppe gefunden'),
                            res && res.found ? 'ok' : 'warn'
                        );
                    } catch (e) {
                        setSummary('Klassenabgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
                    } finally {
                        renderClassesTableFromTextarea();
                    }
                });
                tdAction.appendChild(btnCheck);
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'mini-btn';
                btnDel.textContent = '✕';
                btnDel.title = 'Zeile löschen';
                btnDel.addEventListener('click', () => {
                    const all = getClassesFromTextarea();
                    all.splice(idx, 1);
                    setClassesTextareaFromRows(all);
                    renderClassesTableFromTextarea();
                });
                tdAction.appendChild(btnDel);

                tr.append(tdCode, tdYear, tdName, tdHead, tdEmail, tdMs, tdAction);
                classesTbody.appendChild(tr);
            });
        }

        async function importClassesFromM365Groups() {
            try {
                const searchUnified =
                    graphApi() && typeof graphApi().searchUnifiedGroups === 'function'
                        ? async function (q, token) {
                              const tok = token || (await graphApi().getGraphToken());
                              return graphApi().searchUnifiedGroups(tok, q);
                          }
                        : undefined;
                const picked = await pickEntraGroupsMulti({
                    title: 'Klassengruppen aus Microsoft 365',
                    hint:
                        'Vorschläge mit jg, klasse und class. Ausgewählte Gruppen werden als Klassen übernommen und für den M365-Abgleich verknüpft (typisches Alias: jg2030-1ak).',
                    initialQueries: ['jg', 'klasse', 'class'],
                    searchFn: searchUnified
                });
                if (!picked.length) return;
                const all = getClassesFromTextarea();
                const byCode = new Map();
                all.forEach(function (r) {
                    const c = normCode(r.code);
                    if (c) byCode.set(c, r);
                });
                let added = 0;
                let linked = 0;
                let skipped = 0;
                picked.forEach(function (g) {
                    const derived = deriveClassEntryFromGroup(g);
                    const code = normCode(derived.code);
                    if (!code) {
                        skipped++;
                        return;
                    }
                    let row;
                    if (!byCode.has(code)) {
                        row = {
                            code,
                            year: derived.year || '',
                            name: derived.name || code,
                            headName: '',
                            headEmail: ''
                        };
                        byCode.set(code, row);
                        added++;
                    } else {
                        row = byCode.get(code);
                        if (!normStr(row.year) && derived.year) row.year = derived.year;
                        if (!normStr(row.name) && derived.name) row.name = derived.name;
                    }
                    linkClassGroupMatch(row, g);
                    linked++;
                });
                const next = Array.from(byCode.values()).sort(function (a, b) {
                    return normCode(a.code).localeCompare(normCode(b.code));
                });
                setClassesTextareaFromRows(next);
                renderClassesTableFromTextarea();
                scheduleAutoSave();
                const parts = [];
                if (added) parts.push(added + ' neue Klasse(n)');
                if (linked) parts.push(linked + ' verknüpft');
                if (skipped) parts.push(skipped + ' übersprungen (kein Kürzel)');
                setSummary(
                    'Microsoft 365: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'),
                    added || linked ? 'ok' : 'warn'
                );
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                setSummary('Import aus Microsoft 365: ' + msg, 'warn');
            }
        }

        function buildClassEnrichContext() {
            const teams = [];
            const matchMap = {};
            let prior = new Map();
            try {
                const api = window.ms365AppDataV2;
                if (api && typeof api.getContainer === 'function') {
                    const c = api.getContainer();
                    if (c && c.core && Array.isArray(c.core.classTeams)) {
                        const raw = c.core.classTeams;
                        teams.push(
                            ...(typeof api.normalizeCoreClassTeams === 'function'
                                ? api.normalizeCoreClassTeams(raw)
                                : raw)
                        );
                    }
                    if (c && c.years && c.years.byLabel) {
                        prior = collectPriorClassFieldsByCode(c.years.byLabel);
                    }
                }
                const setup =
                    api && typeof api.getSetup === 'function' ? api.getSetup() : {};
                if (setup && setup.classGroupMatchByKey && typeof setup.classGroupMatchByKey === 'object') {
                    Object.assign(matchMap, setup.classGroupMatchByKey);
                }
            } catch {
                // ignore
            }
            const deriveNick =
                typeof window.ms365DeriveClassStableMailNickname === 'function'
                    ? window.ms365DeriveClassStableMailNickname
                    : null;
            return {
                classTeams: teams,
                classGroupMatchByKey: matchMap,
                priorByCode: prior,
                deriveStableMailNickname: deriveNick
            };
        }

        function applyClassEnrichFromLocalLinks(classes, opts) {
            const options = opts || {};
            const result = enrichClassesFromLinkedGroups(classes || [], buildClassEnrichContext());
            if (!result.changed) return result;
            if (options.persist && typeof save === 'function') {
                const cur = load();
                const saved = save(Object.assign({}, cur, { classes: result.classes }));
                if (taClasses) {
                    taClasses.value = classesToLines(saved.classes || []);
                }
                renderClassesTableFromTextarea();
                if (options.summary !== false && summary) {
                    const st = result.stats || {};
                    setSummary(
                        'Klassenliste ergänzt: ' +
                            (st.yearFilled || 0) +
                            ' Abschlussjahr(e), ' +
                            (st.kvFilled || 0) +
                            ' Klassenvorstand/KV-Daten aus lokalen Verknüpfungen.',
                        'ok'
                    );
                }
                dispatchTenantSettingsChanged(saved, options.reason || 'class-enrich-local');
            }
            return result;
        }

        function buildSubjectEnrichContext() {
            let legacySubjects = [];
            let structureRows = [];
            let appDataContainer = null;
            try {
                const raw = localStorage.getItem('ms365-tenant-settings-v1');
                if (raw) {
                    const o = JSON.parse(raw);
                    if (o && Array.isArray(o.subjects)) legacySubjects = o.subjects.slice();
                }
            } catch {
                // ignore
            }
            try {
                const api = window.ms365AppDataV2;
                if (api && typeof api.getContainer === 'function') {
                    const c = api.getContainer();
                    appDataContainer = c;
                    if (c && c.structure && Array.isArray(c.structure.rows)) {
                        structureRows = c.structure.rows.slice();
                    }
                }
            } catch {
                // ignore
            }
            const cur = load();
            return {
                legacySubjects: legacySubjects,
                structureRows: structureRows,
                arges: (cur && cur.arges) || [],
                appDataContainer: appDataContainer,
                storage: typeof localStorage !== 'undefined' ? localStorage : null
            };
        }

        function applySubjectEnrichFromSources(subjects, opts) {
            const options = opts || {};
            const result = enrichSubjectsFromCatalogSources(subjects || [], buildSubjectEnrichContext());
            if (!result.changed) return result;
            if (options.persist && typeof save === 'function') {
                const cur = load();
                const saved = save(Object.assign({}, cur, { subjects: result.subjects }));
                if (taSubjects) {
                    taSubjects.value = (saved.subjects || [])
                        .map((x) => `${x.code};${x.name || ''}`.trim())
                        .join('\n');
                }
                renderSubjectsTableFromTextarea();
                if (options.summary !== false && summary) {
                    const st = result.stats || {};
                    setSummary(
                        'Fächerliste ergänzt: ' +
                            (st.added || 0) +
                            ' Kürzel neu, ' +
                            (st.namesFilled || 0) +
                            ' Bezeichnungen nachgetragen (Quellen: Unterrichtsteams/Belegung, ARGE, Schulstruktur, ggf. Legacy-Speicher).',
                        'ok'
                    );
                }
                dispatchTenantSettingsChanged(saved, options.reason || 'subject-enrich-local');
            }
            return result;
        }

        function renderFromStorage() {
            let s = load();
            const enriched = applyClassEnrichFromLocalLinks(s.classes || [], {
                persist: true,
                summary: false,
                reason: 'class-enrich-on-load'
            });
            if (enriched && enriched.classes) s = load();
            applySubjectEnrichFromSources(s.subjects || [], {
                persist: true,
                summary: false,
                reason: 'subject-enrich-on-load'
            });
            s = load();
            // Domain in UI-Feld zurückschreiben (wird auch von school-domain.js genutzt)
            try {
                if (schoolNameInput) schoolNameInput.value = normStr(s.schoolName || '');
                if (domainInput) domainInput.value = normStr(s.domain || '');
                if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                    const d = normStr(s.domain || '').replace(/^@+/, '');
                    if (d) window.ms365SetSchoolDomainNoAt(d);
                }
            } catch {
                // ignore
            }
            if (taSubjects) {
                taSubjects.value = (s.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            }
            if (taArges) {
                taArges.value = (s.arges || [])
                    .map((x) => `${x.code};${x.name || ''};${(x.subjects || []).join(',')}`.trim())
                    .join('\n');
            }
            if (taTeachers) {
                taTeachers.value = teachersToLines(s.teachers || []);
            }
            if (taAdminBundle) {
                setAdministrationGroups(adminGroupsFromSettings(s));
            }
            if (taAdmin) {
                taAdmin.value = (s.admin || [])
                    .map((x) => `${x.role || ''};${x.name || ''};${x.email || ''}`.trim())
                    .join('\n');
            }
            if (taAdminRoles) {
                taAdminRoles.value = (s.adminRoles || [])
                    .map((x) => `${x.code || ''};${x.name || ''}`.trim())
                    .join('\n');
            }
            if (selSgaMode) {
                selSgaMode.value = normStr(s.sgaMode || 'group').toLowerCase() === 'distribution' ? 'distribution' : 'group';
            }
            if (taSga) taSga.value = sgaToLines(s.sga || []);
            if (taStudents) {
                taStudents.value = studentsToLines(s.students || []);
            }
            if (taStudentCouncil) taStudentCouncil.value = studentCouncilToLines(s.studentCouncil || []);
            if (taClasses) {
                taClasses.value = (s.classes || [])
                    .map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim())
                    .join('\n');
            }
            renderSubjectsTableFromTextarea();
            renderArgesTableFromTextarea();
            renderTeachersTableFromTextarea();
            renderAdminRolesTableFromTextarea();
            renderAdminTableFromTextarea();
            renderSgaTableFromTextarea();
            renderStudentsTableFromTextarea();
            renderStudentCouncilTableFromTextarea();
            renderClassesTableFromTextarea();
            restoreStoredSchoolWideGroupMatchStatus();
            renderSchoolYearSelectFromV2();
            renderStatusOverview();
            const yLbl = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            setSummary(
                `Aktueller Stand: schulweit ${(s.subjects || []).length} Fächer, ${(s.arges || []).length} ARGEs, ${(s.admin || []).length} Verwaltung, ${(s.sga || []).length} SGA-Einträge, ${(s.teachers || []).length} Lehrkräfte — für Schuljahr ${yLbl}: ${(s.students || []).length} Schüler, ${(s.studentCouncil || []).length} Schülervertretung, ${(s.classes || []).length} Klassen.`,
                'ok'
            );
            dispatchTenantSettingsChanged(s, 'render');
        }

        function renderSchoolYearSelectFromV2() {
            if (!schoolYearSelect) return;
            try {
                if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getContainer !== 'function') {
                    schoolYearSelect.replaceChildren();
                    const o = document.createElement('option');
                    o.value = currentSchoolYearLabel();
                    o.textContent = currentSchoolYearLabel();
                    schoolYearSelect.appendChild(o);
                    schoolYearSelect.value = o.value;
                    return;
                }
                const c = window.ms365AppDataV2.getContainer();
                const cur = c && c.years ? String(c.years.current || '') : '';
                const years = typeof window.ms365AppDataV2.listYears === 'function' ? window.ms365AppDataV2.listYears() : [];
                const list = years.length ? years : (cur ? [cur] : [currentSchoolYearLabel()]);
                schoolYearSelect.replaceChildren();
                list.forEach((y) => {
                    const opt = document.createElement('option');
                    opt.value = String(y);
                    opt.textContent = String(y);
                    schoolYearSelect.appendChild(opt);
                });
                schoolYearSelect.value = cur && list.includes(cur) ? cur : list[0];
            } catch {
                // ignore
            }
        }

        function setCurrentSchoolYearInV2(nextLabel, opts) {
            try {
                if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.setCurrentYear !== 'function') return false;
                window.ms365AppDataV2.setCurrentYear(String(nextLabel || '').trim(), opts || {});
                return true;
            } catch {
                return false;
            }
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
            importSpreadsheetFileToJsonRows(file, onRows, (msg) => setSummary(msg, 'warn'));
        }

        async function importSubjectsRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'fach', 'abk', 'abkuerzung', 'abbreviation']);
                const name = getField(r, ['name', 'fachname', 'bezeichnung', 'subject', 'subjectname']);
                const c = normCode(code);
                if (!c) return;
                out.push({ code: c, name: normStr(name) });
            });
            const ok = await confirmSubjectCatalogIfShrinks(out, 'Fächer-Import (CSV/XLSX)');
            if (!ok) {
                setSummary('Fächer-Import abgebrochen – bestehende Liste unverändert.', 'warn');
                return;
            }
            if (taSubjects) taSubjects.value = out.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            renderSubjectsTableFromTextarea();
            setSummary(`Fächer importiert: ${out.length} (schulweit)`, 'ok');
            scheduleAutoSave();
        }

        function importArgesRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'arge', 'abk', 'abkuerzung']);
                const name = getField(r, ['name', 'bezeichnung', 'titel', 'title']);
                const subs = getField(r, ['faecher', 'fächer', 'subjects', 'fach', 'subjectcodes']);
                const c = normCode(code);
                if (!c) return;
                const subjects = String(subs || '')
                    .split(/[,\s|]+/)
                    .map((x) => normCode(x))
                    .filter(Boolean);
                out.push({ code: c, name: normStr(name), subjects });
            });
            if (taArges) taArges.value = argesToLines(out);
            renderArgesTableFromTextarea();
            setSummary(`ARGEs importiert: ${out.length} (schulweit)`, 'ok');
        }

        function importTeachersRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                const code = getField(r, ['kürzel', 'kuerzel', 'code', 'lehrer', 'abbrev', 'abbreviation']);
                let name = getField(r, ['name', 'lehrername', 'anzeigename', 'displayname']);
                let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                const c = normCode(code);
                if (!c) return;

                // Heuristik für Teillisten: wenn "Name" eigentlich eine E-Mail ist (enthält @),
                // dann korrekt zuordnen statt E-Mail als Name zu speichern.
                const nameNorm = normStr(name);
                const emailNorm = normStr(email).toLowerCase();
                const nameLooksLikeEmail = nameNorm.includes('@');
                const emailLooksLikeEmail = emailNorm.includes('@');

                if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                    email = nameNorm;
                    name = '';
                }

                out.push({ code: c, name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taTeachers) taTeachers.value = teachersToLines(out);
            teachersSortState.key = null;
            teachersSortState.dir = 1;
            renderTeachersTableFromTextarea();
            setSummary(`Lehrkräfte importiert: ${out.length} (schulweit)`, 'ok');
        }

        function importStudentsRows(jsonRows) {
            const prev = getStudentsFromTextarea();
            const sis = window.ms365SchoolSisImport;
            if (sis && typeof sis.importStudentsAndGuardians === 'function') {
                const result = sis.importStudentsAndGuardians({ objectRows: jsonRows, source: 'auto' });
                showSisDiffPreview(prev, result);
                return;
            }
            const out = [];
            (jsonRows || []).forEach((r) => {
                let klasse = getField(r, ['klasse', 'class', 'gruppe', 'group']);
                let name = getField(r, ['name', 'schueler', 'schüler', 'anzeigename', 'displayname']);
                let email = getField(r, ['e-mail', 'email', 'mail', 'upn']);
                if (!klasse && !name && !email) return;

                const nameNorm = normStr(name);
                const emailNorm = normStr(email).toLowerCase();
                const nameLooksLikeEmail = nameNorm.includes('@');
                const emailLooksLikeEmail = emailNorm.includes('@');

                if (nameLooksLikeEmail && (!emailNorm || !emailLooksLikeEmail)) {
                    email = nameNorm;
                    name = '';
                }

                out.push({ klasse: normStr(klasse), name: normStr(name), email: normStr(email).toLowerCase() });
            });
            if (taStudents) taStudents.value = studentsToLines(out);
            studentsSortState.key = null;
            studentsSortState.dir = 1;
            renderStudentsTableFromTextarea();
            const ySt = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            setSummary(`Schüler importiert: ${out.length} (Schuljahr ${ySt})`, 'ok');
            renderStudentLifecyclePanel(prev, out);
        }

        function importClassesFromWebuntisPdfFile(file) {
            const wu = window.ms365WebuntisExportImport;
            const pdfApi = window.ms365PdfText;
            if (!wu || typeof wu.importClassesFromWebuntisPdf !== 'function') {
                setSummary('WebUntis-Klassenimport: Modul fehlt – Seite neu laden.', 'warn');
                return;
            }
            if (!pdfApi || typeof pdfApi.extractPdfFile !== 'function') {
                setSummary('PDF-Modul fehlt – Seite neu laden.', 'warn');
                return;
            }
            const skipEl = document.getElementById('tenantClassesSkipNoKv') || document.getElementById('swClassesSkipNoKv');
            const skipWithoutTeacher = skipEl ? !!skipEl.checked : true;
            const teachers = getTeachersFromTextarea();
            setSummary('Klassen-PDF wird gelesen …', 'ok');
            pdfApi
                .extractPdfFile(file)
                .then(function (content) {
                    const result = wu.importClassesFromWebuntisPdf({
                        text: content.text,
                        words: content.words,
                        teachers: teachers,
                        skipWithoutTeacher: skipWithoutTeacher,
                        inferYear: true
                    });
                    if (!result.classes.length) {
                        setSummary('Keine Klassen im PDF erkannt (Kurzname / Klassenlehrkraft prüfen).', 'warn');
                        return;
                    }
                    if (taClasses) taClasses.value = result.lines;
                    renderClassesTableFromTextarea();
                    const bits = [
                        result.meta.classCount + ' Klassen',
                        result.meta.withTeacher + ' mit KV-Kürzel',
                        result.meta.matched + ' KV aus Lehrerliste'
                    ];
                    if (result.meta.missingTeacher) bits.push(result.meta.missingTeacher + ' Kürzel ohne Lehrerliste');
                    if (result.schoolYear && result.schoolYear.label) bits.push('Schuljahr ' + result.schoolYear.label);
                    setSummary('WebUntis-Klassen importiert: ' + bits.join(' · '), 'ok');
                    scheduleAutoSave();
                })
                .catch(function (err) {
                    setSummary('Klassen-PDF: ' + (err && err.message ? err.message : String(err)), 'warn');
                });
        }

        function importClassesRows(jsonRows) {
            const out = [];
            (jsonRows || []).forEach((r) => {
                let code = getField(r, ['abkürzung', 'abkuerzung', 'abk', 'kuerzel', 'kürzel', 'code', 'klasseabk', 'classcode', 'kurzname']);
                let year = getField(r, ['abschlussjahr', 'abschluss', 'year', 'graduationyear']);
                let name = getField(r, ['klasse', 'class', 'name', 'bezeichnung', 'classname', 'langname']);
                let headName = getField(r, ['klassenvorstand', 'klassenvorstandname', 'kv', 'kvname', 'vorstand', 'head', 'headname']);
                let headEmail = getField(r, ['klassenvorstandemail', 'kvemail', 'e-mail', 'email', 'mail', 'upn', 'heademail']);
                const headCode = getField(r, ['klassenlehrkraft', 'lehrkraft', 'kvkuerzel', 'headcode']);
                if (!code && !year && !name && !headName && !headEmail && !headCode) return;

                // Heuristik: falls "Klassenvorstand" eigentlich E-Mail ist
                const hn = normStr(headName);
                const he = normStr(headEmail).toLowerCase();
                if (hn.includes('@') && (!he || !he.includes('@'))) {
                    headEmail = hn;
                    headName = '';
                }

                let resolvedName = normStr(headName);
                let resolvedEmail = normStr(headEmail).toLowerCase();
                if (headCode && (!resolvedName || !resolvedEmail)) {
                    const teachers = getTeachersFromTextarea();
                    const hit = teachers.find((t) => normCode(t.code) === normCode(headCode));
                    if (hit) {
                        if (!resolvedName) resolvedName = normStr(hit.name);
                        if (!resolvedEmail) resolvedEmail = normStr(hit.email).toLowerCase();
                    } else if (!resolvedName) {
                        resolvedName = normCode(headCode);
                    }
                }

                out.push({
                    code: normCode(code),
                    year: /^\d{4}$/.test(normStr(year)) ? normStr(year) : '',
                    name: normStr(name),
                    headName: resolvedName,
                    headEmail: resolvedEmail
                });
            });
            if (taClasses) taClasses.value = classesToLines(out);
            renderClassesTableFromTextarea();
            const yCl = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
            setSummary(`Klassen importiert: ${out.length} (Schuljahr ${yCl})`, 'ok');
        }

        if (btnSave) {
            btnSave.addEventListener('click', async () => {
                const prevLoaded = load();
                const prevStudents = (prevLoaded && prevLoaded.students) || [];
                const subjects = typeof parseLinesToSubjects === 'function' ? parseLinesToSubjects(taSubjects ? taSubjects.value : '') : [];
                const okSubjects = await confirmSubjectCatalogIfShrinks(
                    subjects,
                    'Speichern der Stammdaten'
                );
                if (!okSubjects) {
                    setSummary('Speichern abgebrochen – Fächerliste unverändert.', 'warn');
                    return;
                }
                const teachers = typeof parseLinesToTeachers === 'function' ? parseLinesToTeachers(taTeachers ? taTeachers.value : '') : [];
                const admin = getAdminFromTextarea();
                const adminRoles = getAdminRolesFromTextarea();
                const sga = typeof parseLinesToSga === 'function' ? parseLinesToSga(taSga ? taSga.value : '') : [];
                const sgaMode = selSgaMode ? normStr(selSgaMode.value || 'group').toLowerCase() : 'group';
                const students = typeof parseLinesToStudents === 'function' ? parseLinesToStudents(taStudents ? taStudents.value : '') : [];
                const studentCouncil =
                    typeof parseLinesToStudentCouncil === 'function'
                        ? parseLinesToStudentCouncil(taStudentCouncil ? taStudentCouncil.value : '')
                        : [];
                const classes = typeof parseLinesToClasses === 'function' ? parseLinesToClasses(taClasses ? taClasses.value : '') : [];
                const schoolName = schoolNameInput ? normStr(schoolNameInput.value || '') : '';
                const domain =
                    typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
                const arges = typeof parseLinesToArges === 'function' ? parseLinesToArges(taArges ? taArges.value : '') : [];
                const administration = getAdministrationEntries();
                const saved = save({ schoolName, domain, subjects, arges, teachers, administration, admin, adminRoles, sgaMode, sga, students, studentCouncil, classes });
                const ySave = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
                setSummary(
                    `Gespeichert: schulweit ${(saved.subjects || []).length} Fächer, ${(saved.arges || []).length} ARGEs, ${(saved.admin || []).length} Verwaltung, ${(saved.sga || []).length} SGA-Einträge, ${(saved.teachers || []).length} Lehrkräfte — für Schuljahr ${ySave}: ${(saved.students || []).length} Schüler, ${(saved.studentCouncil || []).length} Schülervertretung, ${(saved.classes || []).length} Klassen.`,
                    'ok'
                );
                renderSubjectsTableFromTextarea();
                renderArgesTableFromTextarea();
                renderTeachersTableFromTextarea();
                renderAdminRolesTableFromTextarea();
                renderAdminTableFromTextarea();
                renderSgaTableFromTextarea();
                renderStudentsTableFromTextarea();
                renderStudentCouncilTableFromTextarea();
                renderClassesTableFromTextarea();
                renderStatusOverview();
                renderStudentLifecyclePanel(prevStudents, students);
                dispatchTenantSettingsChanged(saved, 'manual-save');
            });
        }

        // Struktur -> Listen (Writeback)
        function mergeIntoSubjects(existing, codesToEnsure) {
            const out = (existing || []).slice();
            const seen = new Set(out.map((s) => String(s.code || '').toUpperCase()).filter(Boolean));
            (codesToEnsure || []).forEach((c) => {
                const code = normCode(c);
                if (!code || seen.has(code)) return;
                seen.add(code);
                out.push({ code, name: '' });
            });
            out.sort((a, b) => normCode(a.code).localeCompare(normCode(b.code)));
            return out;
        }

        function mergeIntoClasses(existing, codesToEnsure) {
            const out = (existing || []).slice();
            const seen = new Set(out.map((c) => normCode(c.code || c.name || '')).filter(Boolean));
            (codesToEnsure || []).forEach((c) => {
                const code = normCode(c);
                if (!code || seen.has(code)) return;
                seen.add(code);
                out.push({ code, name: `Klasse ${code}`, year: '', headName: '', headEmail: '' });
            });
            out.sort((a, b) => normCode(a.code).localeCompare(normCode(b.code)));
            return out;
        }

        function mergeIntoArges(existing, codesToEnsure) {
            const out = (existing || []).slice();
            const seen = new Set(out.map((a) => normCode(a.code || '')).filter(Boolean));
            (codesToEnsure || []).forEach((c) => {
                const code = normCode(c);
                if (!code || seen.has(code)) return;
                seen.add(code);
                out.push({ code, name: code, subjects: [] });
            });
            out.sort((a, b) => normCode(a.code).localeCompare(normCode(b.code)));
            return out;
        }

        function applyWriteback(detail) {
            if (!detail || !detail.writeback) return;
            if (typeof load !== 'function' || typeof save !== 'function') return;
            const current = load();
            const next = Object.assign({}, current);
            if (detail.writeback.subjectCodes) {
                next.subjects = mergeIntoSubjects(current.subjects || [], detail.writeback.subjectCodes);
            }
            if (detail.writeback.argeCodes) {
                next.arges = mergeIntoArges(current.arges || [], detail.writeback.argeCodes);
            }
            if (detail.writeback.classCodes) {
                next.classes = mergeIntoClasses(current.classes || [], detail.writeback.classCodes);
            }
            __syncGuard++;
            const saved = save(next);
            // UI aktualisieren
            try {
                if (taSubjects) taSubjects.value = subjectsToLines(saved.subjects || []);
                if (taArges) taArges.value = argesToLines(saved.arges || []);
                if (taClasses) taClasses.value = classesToLines(saved.classes || []);
                renderSubjectsTableFromTextarea();
                renderArgesTableFromTextarea();
                renderClassesTableFromTextarea();
                setSummary('Listen wurden aus der Struktur ergänzt.', 'ok');
            } catch {
                // ignore
            }
            __syncGuard--;
            dispatchTenantSettingsChanged(saved, 'writeback');
        }

        try {
            window.addEventListener('ms365-structure-changed', (ev) => {
                if (__syncGuard) return;
                applyWriteback(ev && ev.detail ? ev.detail : null);
                try {
                    if (typeof window.__ms365TenantStepsRefreshMatch === 'function') window.__ms365TenantStepsRefreshMatch();
                } catch {
                    // ignore
                }
            });
            window.addEventListener('ms365-match-links-changed', () => {
                try {
                    if (typeof window.__ms365TenantStepsRefreshMatch === 'function') window.__ms365TenantStepsRefreshMatch();
                } catch {
                    // ignore
                }
            });
        } catch {
            // ignore
        }

        if (fileSubjects) {
            fileSubjects.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importSubjectsRows(rows));
                fileSubjects.value = '';
            });
        }
        if (fileArges) {
            fileArges.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                importFileToRows(f, (rows) => importArgesRows(rows));
                fileArges.value = '';
            });
        }
        if (btnArgesImportM365) {
            btnArgesImportM365.addEventListener('click', () => {
                importArgesFromM365Groups();
            });
        }
        if (btnSubjectsImportM365) {
            btnSubjectsImportM365.addEventListener('click', () => {
                importSubjectsFromM365Groups();
            });
        }
        if (btnClassesImportM365) {
            btnClassesImportM365.addEventListener('click', () => {
                importClassesFromM365Groups();
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
                const files = e.target.files;
                const sourceEl = document.getElementById('tenantStudentsImportSource');
                const sourceHint = sourceEl ? String(sourceEl.value || 'auto') : 'auto';
                if (window.ms365StudentListImport && typeof window.ms365StudentListImport.importFile === 'function') {
                    window.ms365StudentListImport.importFile(
                        files,
                        (lines, result) => {
                            const prev = getStudentsFromTextarea();
                            showSisDiffPreview(prev, result || { lines: lines, records: [], source: sourceHint });
                            const meta = result && result.meta ? result.meta : {};
                            const bits = [];
                            if (meta.emailsGenerated) bits.push(meta.emailsGenerated + ' Mails vorgeschlagen');
                            if (meta.unmatchedGuardians) bits.push(meta.unmatchedGuardians + ' Eltern ohne Zuordnung');
                            if (meta.emailHint) bits.push(meta.emailHint);
                            if (bits.length) setSummary('WebUntis/SIS: ' + bits.join(' · '), 'ok');
                        },
                        (msg) => setSummary(msg, 'warn'),
                        sourceHint,
                        getStudentsFromTextarea()
                    );
                } else {
                    const f = files && files[0];
                    importFileToRows(f, (rows) => importStudentsRows(rows));
                }
                fileStudents.value = '';
            });
        }
        const btnSisMerge = document.getElementById('tenantSisDiffMerge');
        const btnSisReplace = document.getElementById('tenantSisDiffReplace');
        const btnSisCancel = document.getElementById('tenantSisDiffCancel');
        if (btnSisMerge) btnSisMerge.addEventListener('click', () => applyPendingSis('merge'));
        if (btnSisReplace) btnSisReplace.addEventListener('click', () => applyPendingSis('replace'));
        if (btnSisCancel) btnSisCancel.addEventListener('click', hideSisDiff);

        function renderActionLog() {
            const ul = document.getElementById('tenantActionLogList');
            if (!ul) return;
            ul.innerHTML = '';
            const rows =
                window.ms365ActionLog && typeof window.ms365ActionLog.list === 'function'
                    ? window.ms365ActionLog.list(40)
                    : [];
            if (!rows.length) {
                const li = document.createElement('li');
                li.textContent = 'Noch keine Einträge.';
                ul.appendChild(li);
                return;
            }
            rows.forEach(function (row) {
                const li = document.createElement('li');
                const when = row.at ? String(row.at).replace('T', ' ').slice(0, 19) : '';
                li.textContent =
                    (when ? when + ' · ' : '') +
                    (row.tool || 'app') +
                    ': ' +
                    (row.summary || row.action) +
                    (row.result === 'error' ? ' (Fehler)' : '');
                ul.appendChild(li);
            });
        }
        const btnLogRefresh = document.getElementById('tenantActionLogRefresh');
        const btnLogExport = document.getElementById('tenantActionLogExport');
        const btnLogClear = document.getElementById('tenantActionLogClear');
        if (btnLogRefresh) btnLogRefresh.addEventListener('click', renderActionLog);
        if (btnLogExport) {
            btnLogExport.addEventListener('click', function () {
                if (!window.ms365ActionLog || typeof window.ms365ActionLog.exportJson !== 'function') return;
                const blob = new Blob([window.ms365ActionLog.exportJson()], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'ms365-aktionsprotokoll.json';
                document.body.appendChild(a);
                a.click();
                a.remove();
                setTimeout(function () {
                    URL.revokeObjectURL(url);
                }, 250);
            });
        }
        if (btnLogClear) {
            btnLogClear.addEventListener('click', function () {
                if (window.ms365ActionLog) window.ms365ActionLog.clear();
                renderActionLog();
            });
        }
        renderActionLog();

        if (fileClasses) {
            fileClasses.addEventListener('change', (e) => {
                const f = e.target.files && e.target.files[0];
                fileClasses.value = '';
                if (!f) return;
                const name = String(f.name || '').toLowerCase();
                if (name.endsWith('.pdf')) {
                    importClassesFromWebuntisPdfFile(f);
                    return;
                }
                importFileToRows(f, (rows) => importClassesRows(rows));
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
        if (btnArgesTpl) {
            btnArgesTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'ARGE-Liste-Vorlage.xlsx',
                    [
                        ['Kürzel', 'Name', 'Fächer'],
                        ['SPRACHEN', 'Sprachen', 'D,E,FS2'],
                        ['NAWI', 'Naturwissenschaften', 'BIO,CH,PH']
                    ],
                    'ARGEs'
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
                const ok =
                    window.ms365StudentListImport && typeof window.ms365StudentListImport.downloadTemplate === 'function'
                        ? window.ms365StudentListImport.downloadTemplate()
                        : false;
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }
        const btnStudentsCsvTpl = document.getElementById('tenantStudentsTemplateCsv');
        if (btnStudentsCsvTpl) {
            btnStudentsCsvTpl.addEventListener('click', () => {
                const ok =
                    window.ms365StudentListImport && typeof window.ms365StudentListImport.downloadCsvTemplate === 'function'
                        ? window.ms365StudentListImport.downloadCsvTemplate()
                        : false;
                if (!ok) setSummary('CSV-Vorlage konnte nicht erzeugt werden.', 'warn');
            });
        }
        if (btnClassesTpl) {
            btnClassesTpl.addEventListener('click', () => {
                const ok = downloadXlsxTemplate(
                    'Klassenliste-Vorlage.xlsx',
                    [
                        ['Abkürzung', 'Abschlussjahr', 'Klasse', 'Klassenvorstand', 'E-Mail'],
                        ['1AK', '2030', '1A-Klasse', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['2BK', '2029', '2B-Klasse', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ],
                    'Klassen'
                );
                if (!ok) setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
            });
        }

        if (btnReload) {
            btnReload.addEventListener('click', () => renderFromStorage());
        }

        if (schoolYearSelect && !schoolYearSelect.dataset.bound) {
            schoolYearSelect.dataset.bound = '1';
            schoolYearSelect.addEventListener('change', () => {
                const y = String(schoolYearSelect.value || '').trim();
                if (!y) return;
                setCurrentSchoolYearInV2(y);
                patchMatchedGroupId('studentCouncil', null);
                renderFromStorage();
                setSummary('Schuljahr gewechselt: ' + y + ' — Schüler-, Schülervertretungs- und Klassenlisten beziehen sich nun auf dieses Jahr.', 'ok');
            });
        }
        if (schoolYearAddBtn && !schoolYearAddBtn.dataset.bound) {
            schoolYearAddBtn.dataset.bound = '1';
            schoolYearAddBtn.addEventListener('click', () => {
                void (async () => {
                    const cur = schoolYearSelect ? String(schoolYearSelect.value || '').trim() : '';
                    const suggest = (function () {
                        const m = cur.match(/^(\d{4})\s*\/\s*(\d{2}|\d{4})/);
                        if (!m) return '';
                        const y = parseInt(m[1], 10);
                        if (!isFinite(y)) return '';
                        return String(y + 1) + '/' + String(y + 2).slice(2);
                    })();
                    const next = await dlgPrompt('Neues Schuljahr (z. B. 2027/28)', suggest || currentSchoolYearLabel(), {
                        title: 'Schuljahr',
                        inputLabel: 'Bezeichnung'
                    });
                    if (next == null || !normStr(next)) return;
                    const copy = await dlgConfirm('Schüler & Klassen aus dem aktuellen Schuljahr übernehmen?', {
                        title: 'Schuljahr',
                        okText: 'Ja, übernehmen',
                        cancelText: 'Nein'
                    });
                    setCurrentSchoolYearInV2(next, copy && cur ? { copyFrom: cur } : {});
                    renderFromStorage();
                    if (schoolYearSelect) schoolYearSelect.value = String(next).trim();
                    setSummary('Neues Schuljahr angelegt: ' + String(next).trim(), 'ok');
                })();
            });
        }

        if (btnExport) {
            btnExport.addEventListener('click', () => {
                try {
                    if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.downloadBackup === 'function') {
                        window.ms365BrowserBackup.downloadBackup();
                        return;
                    }
                    if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.exportJson === 'function') {
                        downloadJson('ms365-schooltool-data-v2.json', window.ms365AppDataV2.exportJson());
                        return;
                    }
                } catch {
                    // ignore
                }
                downloadJson('schule-einstellungen.json', load());
            });
        }
        if (btnExportHeader) {
            btnExportHeader.addEventListener('click', () => {
                try {
                    if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.downloadBackup === 'function') {
                        window.ms365BrowserBackup.downloadBackup();
                        return;
                    }
                    if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.exportJson === 'function') {
                        downloadJson('ms365-schooltool-data-v2.json', window.ms365AppDataV2.exportJson());
                        return;
                    }
                } catch {
                    // ignore
                }
                downloadJson('schule-einstellungen.json', load());
            });
        }

        if (btnClear) {
            btnClear.addEventListener('click', () => {
                try {
                    localStorage.removeItem('ms365-tenant-settings-v1');
                } catch {
                    // ignore
                }
                // UI/Domain wieder auf Standard zurücksetzen
                try {
                    if (schoolNameInput) schoolNameInput.value = '';
                    const domainInput = document.getElementById('schoolEmailDomain');
                    if (domainInput) domainInput.value = 'ms365.schule';
                    if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                        window.ms365SetSchoolDomainNoAt('ms365.schule');
                    }
                } catch {
                    // ignore
                }
                if (taSubjects) taSubjects.value = '';
                if (taArges) taArges.value = '';
                if (taTeachers) taTeachers.value = '';
                if (taAdminBundle) taAdminBundle.value = '';
                if (taAdmin) taAdmin.value = '';
                if (taAdminRoles) taAdminRoles.value = '';
                if (selSgaMode) selSgaMode.value = 'group';
                if (taSga) taSga.value = '';
                if (taStudents) taStudents.value = '';
                if (taStudentCouncil) taStudentCouncil.value = '';
                if (taClasses) taClasses.value = '';
                renderSubjectsTableFromTextarea();
                renderArgesTableFromTextarea();
                renderTeachersTableFromTextarea();
                renderAdminRolesTableFromTextarea();
                renderAdminTableFromTextarea();
                renderSgaTableFromTextarea();
                renderStudentsTableFromTextarea();
                renderStudentCouncilTableFromTextarea();
                renderClassesTableFromTextarea();
                setSummary('Stammdaten gelöscht (nur lokaler Browser-Speicher).', 'warn');
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
                    if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.isBackupPayload === 'function' && window.ms365BrowserBackup.isBackupPayload(obj)) {
                        let backupMsg =
                            'Dieses Browser-Backup ersetzt alle lokalen Schuldaten und wiederherstellbaren Werkzeug-Zwischenstände in diesem Browser. Microsoft-Anmeldung und PIN-Freischaltungen bleiben unberührt. Fortfahren?';
                        const backupSubjects = subjectsFromImportPayload(obj);
                        if (backupSubjects.length) {
                            const shrinkMsg = formatSubjectCatalogShrinkConfirmDe(
                                subjectCatalogShrinkReport(
                                    typeof load === 'function' ? load().subjects || [] : [],
                                    backupSubjects
                                ),
                                ''
                            );
                            if (shrinkMsg) backupMsg += '\n\n' + shrinkMsg;
                        }
                        const ok = await dlgConfirm(backupMsg, {
                            title: 'Backup importieren',
                            okText: 'Importieren',
                            cancelText: 'Abbrechen',
                            danger: true
                        });
                        if (!ok) return;
                        window.ms365BrowserBackup.applyBackup(obj);
                        location.reload();
                        return;
                    }
                    const isV2 =
                        obj.version >= 2 && obj.core && obj.structure && obj.match;
                    const importSubjects = subjectsFromImportPayload(obj);
                    if (importSubjects.length) {
                        const okSub = await confirmSubjectCatalogIfShrinks(
                            importSubjects,
                            isV2 ? 'JSON-Import (Schuldaten v2)' : 'JSON-Import (Stammdaten)'
                        );
                        if (!okSub) {
                            setSummary('Import abgebrochen – Fächerliste würde verkleinert worden sein.', 'warn');
                            return;
                        }
                    }
                    try {
                        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.importJson === 'function') {
                            window.ms365AppDataV2.importJson(obj);
                        }
                    } catch (errImport) {
                        setSummary('Import fehlgeschlagen: ' + (errImport?.message || String(errImport)), 'warn');
                        return;
                    }
                    const saved = isV2 ? load() : save(obj);
                    if (schoolNameInput) schoolNameInput.value = normStr(saved.schoolName || '');
                    if (taSubjects) taSubjects.value = (saved.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
                    if (taArges) taArges.value = (saved.arges || []).map((x) => `${x.code};${x.name || ''};${(x.subjects || []).join(',')}`.trim()).join('\n');
                    if (taTeachers) taTeachers.value = teachersToLines(saved.teachers || []);
                    if (taAdminBundle) setAdministrationGroups(adminGroupsFromSettings(saved));
                    if (taAdmin) taAdmin.value = (saved.admin || []).map((x) => `${x.role || ''};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                    if (taAdminRoles) taAdminRoles.value = (saved.adminRoles || []).map((x) => `${x.code || ''};${x.name || ''}`.trim()).join('\n');
                    if (selSgaMode) selSgaMode.value = normStr(saved.sgaMode || 'group').toLowerCase() === 'distribution' ? 'distribution' : 'group';
                    if (taSga) taSga.value = sgaToLines(saved.sga || []);
                    if (taStudents) taStudents.value = studentsToLines(saved.students || []);
                    if (taStudentCouncil) taStudentCouncil.value = studentCouncilToLines(saved.studentCouncil || []);
                    if (taClasses) taClasses.value = (saved.classes || []).map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim()).join('\n');
                    renderSubjectsTableFromTextarea();
                    renderArgesTableFromTextarea();
                    renderTeachersTableFromTextarea();
                    renderAdminRolesTableFromTextarea();
                    renderAdminTableFromTextarea();
                    renderSgaTableFromTextarea();
                    renderStudentsTableFromTextarea();
                    renderStudentCouncilTableFromTextarea();
                    renderClassesTableFromTextarea();
                    const yImp = getDisplayedSchoolYearLabel() || currentSchoolYearLabel();
                    setSummary(
                        `Import OK: schulweit ${(saved.subjects || []).length} Fächer, ${(saved.arges || []).length} ARGEs, ${(saved.admin || []).length} Verwaltung, ${(saved.sga || []).length} SGA-Einträge, ${(saved.teachers || []).length} Lehrkräfte — für Schuljahr ${yImp}: ${(saved.students || []).length} Schüler, ${(saved.studentCouncil || []).length} Schülervertretung, ${(saved.classes || []).length} Klassen.`,
                        'ok'
                    );
                } catch (err) {
                    setSummary('Import fehlgeschlagen: ' + (err?.message || String(err)), 'warn');
                } finally {
                    fileImport.value = '';
                }
            });
        }

        if (domainInput) {
            domainInput.addEventListener('input', () => scheduleAutoSave());
            domainInput.addEventListener('change', () => scheduleAutoSave());
        }
        if (schoolNameInput) {
            schoolNameInput.addEventListener('input', () => {
                clearStoredSchoolWideGroupMatches();
                scheduleAutoSave();
            });
            schoolNameInput.addEventListener('change', () => {
                clearStoredSchoolWideGroupMatches();
                scheduleAutoSave();
            });
        }
        // Kein Standard-Abschlussjahr mehr in den Stammdaten
        if (taSubjects) {
            taSubjects.addEventListener('input', () => {
                subjectsSortState.key = null;
                subjectsSortState.dir = 1;
                renderSubjectsTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taArges) {
            taArges.addEventListener('input', () => {
                argesSortState.key = null;
                argesSortState.dir = 1;
                renderArgesTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (btnAddSubjectRow) {
            btnAddSubjectRow.addEventListener('click', () => {
                const all = getSubjectsFromTextarea();
                all.push({ code: '', name: '' });
                setSubjectsTextareaFromRows(all);
                renderSubjectsTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnEnrichSubjects && !btnEnrichSubjects.dataset.tenantEnrichSubjectsBound) {
            btnEnrichSubjects.dataset.tenantEnrichSubjectsBound = '1';
            btnEnrichSubjects.addEventListener('click', () => {
                const rows = getSubjectsFromTextarea();
                const result = applySubjectEnrichFromSources(rows, {
                    persist: true,
                    summary: true,
                    reason: 'subject-enrich-manual'
                });
                if (!result.changed) {
                    setSummary(
                        'Keine zusätzlichen Fächer in ARGE, Schulstruktur oder Legacy-Speicher gefunden – Liste unverändert.',
                        'warn'
                    );
                }
                renderStatusOverview();
            });
        }

        if (btnAddArgeRow) {
            btnAddArgeRow.addEventListener('click', () => {
                const all = getArgesFromTextarea();
                all.push({ code: '', name: '', subjects: [] });
                setArgesTextareaFromRows(all);
                renderArgesTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taTeachers) {
            taTeachers.addEventListener('input', () => {
                teachersSortState.key = null;
                teachersSortState.dir = 1;
                renderTeachersTableFromTextarea();
            });
            taTeachers.addEventListener('input', () => scheduleAutoSave());
        }
        if (teachersTable && !teachersTable.dataset.teachersSortBound) {
            teachersTable.dataset.teachersSortBound = '1';
            teachersTable.querySelectorAll('th.is-sortable').forEach((th) => {
                th.addEventListener('click', () => applyTeachersSort(th.dataset.sortKey));
                th.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        applyTeachersSort(th.dataset.sortKey);
                    }
                });
            });
        }
        bindTenantTableSort(subjectsTable, subjectsSortState, ['code', 'name'], applySubjectsSort);
        bindTenantTableSort(argesTable, argesSortState, ['code', 'name', 'subjects'], applyArgesSort);
        bindTenantTableSort(classesTable, classesSortState, ['code', 'year', 'name', 'headName', 'headEmail'], applyClassesSort);
        bindTenantTableSort(sgaTable, sgaSortState, ['scope', 'name', 'email'], applySgaSort);
        bindTenantTableSort(
            studentCouncilTable,
            studentCouncilSortState,
            ['klasse', 'name', 'email'],
            applyStudentCouncilSort
        );
        bindTenantTableSort(
            adminUnifiedTable,
            adminUnifiedSortState,
            ['name', 'code', 'personName', 'email'],
            applyAdminUnifiedSort
        );
        if (btnAddTeacherRow) {
            btnAddTeacherRow.addEventListener('click', () => {
                const all = getTeachersFromTextarea();
                all.push({ code: '', name: '', email: '' });
                setTeachersTextareaFromRows(all);
                renderTeachersTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnTeachersExportCsv) {
            btnTeachersExportCsv.addEventListener('click', () => {
                const api = window.ms365TeacherListImport;
                const rows = getTeachersFromTextarea();
                if (!rows.length) {
                    setSummary('Lehrerliste ist leer – nichts zu exportieren.', 'warn');
                    return;
                }
                if (!api || typeof api.exportCsv !== 'function' || !api.exportCsv(rows)) {
                    setSummary('CSV-Export fehlgeschlagen.', 'warn');
                    return;
                }
                setSummary('Lehrerliste als CSV exportiert (' + rows.length + ').', 'ok');
            });
        }
        if (btnTeachersExportXlsx) {
            btnTeachersExportXlsx.addEventListener('click', () => {
                const api = window.ms365TeacherListImport;
                const rows = getTeachersFromTextarea();
                if (!rows.length) {
                    setSummary('Lehrerliste ist leer – nichts zu exportieren.', 'warn');
                    return;
                }
                if (!api || (typeof api.isXlsxReady === 'function' && !api.isXlsxReady())) {
                    setSummary('Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
                    return;
                }
                if (!api || typeof api.exportXlsx !== 'function' || !api.exportXlsx(rows)) {
                    setSummary('XLSX-Export fehlgeschlagen.', 'warn');
                    return;
                }
                setSummary('Lehrerliste als XLSX exportiert (' + rows.length + ').', 'ok');
            });
        }
        if (btnVerifyTeachersGraph && !btnVerifyTeachersGraph.dataset.tenantVerifyTeachersBound) {
            btnVerifyTeachersGraph.dataset.tenantVerifyTeachersBound = '1';
            btnVerifyTeachersGraph.addEventListener('click', async () => {
                await runVerifyGraphDirectoryRows(
                    getTeachersFromTextarea(),
                    function (r) {
                        return r.email;
                    },
                    'Lehrkräfte',
                    btnVerifyTeachersGraph,
                    function () {
                        renderTeachersTableFromTextarea();
                    }
                );
            });
        }
        if (taAdminBundle) {
            taAdminBundle.addEventListener('input', () => {
                adminUnifiedSortState.key = null;
                adminUnifiedSortState.dir = 1;
                renderAdminRolesTableFromTextarea();
                renderAdminTableFromTextarea();
                renderAdminUnifiedTableFromBundle();
                scheduleAutoSave();
            });
        }
        if (taAdmin) {
            taAdmin.addEventListener('input', () => renderAdminTableFromTextarea());
            taAdmin.addEventListener('input', () => scheduleAutoSave());
        }
        if (taAdminRoles) {
            taAdminRoles.addEventListener('input', () => renderAdminRolesTableFromTextarea());
            taAdminRoles.addEventListener('input', () => renderAdminTableFromTextarea());
            taAdminRoles.addEventListener('input', () => scheduleAutoSave());
        }
        if (btnAddAdminRow) {
            btnAddAdminRow.addEventListener('click', () => {
                const rows = groupsToDisplayRows(getAdministrationGroups());
                const newName = '';
                rows.push({ code: '', name: newName, personName: '', email: '' });
                setAdministrationGroups(displayRowsToGroups(rows));
                renderAdminTableFromTextarea();
                scheduleAutoSave();
                // Letzte Zeile sofort in Editiermodus (Bezeichnung-Zelle)
                if (adminUnifiedTbody) {
                    const lastTr = adminUnifiedTbody.lastElementChild;
                    if (lastTr) {
                        const tdLabel = lastTr.cells && lastTr.cells[0];
                        if (tdLabel) {
                            const newIdx = rows.length - 1;
                            startCellEdit(tdLabel, newName, (next, meta) => {
                                const all = groupsToDisplayRows(getAdministrationGroups());
                                if (!all[newIdx]) return renderAdminUnifiedTableFromBundle();
                                all[newIdx].name = meta && meta.cancelled ? '' : normStr(next);
                                if (!meta || !meta.cancelled) {
                                    all[newIdx].code =
                                        typeof window.ms365TenantSettingsAdminRoleCodeFromName === 'function'
                                            ? window.ms365TenantSettingsAdminRoleCodeFromName(all[newIdx].name)
                                            : '';
                                }
                                setAdministrationGroups(displayRowsToGroups(all));
                                renderAdminUnifiedTableFromBundle();
                                scheduleAutoSave();
                            });
                        }
                    }
                }
            });
        }
        if (btnAdminRolesDefaults) {
            btnAdminRolesDefaults.addEventListener('click', () => {
                const defaults =
                    typeof window.ms365TenantSettingsDefaultAdminRoles === 'function'
                        ? window.ms365TenantSettingsDefaultAdminRoles()
                        : [];
                const rows = groupsToDisplayRows(getAdministrationGroups());
                const seen = new Set(
                    rows.map(function (row) {
                        return normStr(row.name).toLowerCase();
                    })
                );
                defaults.forEach(function (d) {
                    const name = normStr(d && d.name);
                    const key = name.toLowerCase();
                    if (!name || seen.has(key)) return;
                    seen.add(key);
                    rows.push({ code: normCode(d && d.code), name: name, personName: '', email: '' });
                });
                setAdministrationGroups(displayRowsToGroups(rows));
                renderAdminTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnVerifyVerwaltungGraph && !btnVerifyVerwaltungGraph.dataset.tenantVerifyVerwaltungBound) {
            btnVerifyVerwaltungGraph.dataset.tenantVerifyVerwaltungBound = '1';
            btnVerifyVerwaltungGraph.addEventListener('click', async () => {
                await runVerifyVerwaltungGraphBulk();
            });
        }
        if (selSgaMode) {
            selSgaMode.addEventListener('change', () => {
                clearStoredSchoolWideGroupMatches();
                scheduleAutoSave();
            });
        }
        if (taSga) {
            taSga.addEventListener('input', () => {
                sgaSortState.key = null;
                sgaSortState.dir = 1;
                renderSgaTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnAddSgaRow) {
            btnAddSgaRow.addEventListener('click', () => {
                const all = getSgaFromTextarea();
                all.push({ scope: 'teacher', name: '', email: '' });
                setSgaTextareaFromRows(all);
                renderSgaTableFromTextarea();
                scheduleAutoSave();
            });
        }

        if (taStudents) {
            let studentsTableRenderTimer = null;
            taStudents.addEventListener('input', () => {
                studentsSortState.key = null;
                studentsSortState.dir = 1;
                if (studentsTableRenderTimer) clearTimeout(studentsTableRenderTimer);
                studentsTableRenderTimer = setTimeout(() => {
                    studentsTableRenderTimer = null;
                    renderStudentsTableFromTextarea();
                }, 250);
            });
            taStudents.addEventListener('input', () => scheduleAutoSave());
        }
        if (studentsClassFilter && !studentsClassFilter.dataset.studentsFilterBound) {
            studentsClassFilter.dataset.studentsFilterBound = '1';
            studentsClassFilter.addEventListener('change', () => {
                renderStudentsTableFromTextarea();
            });
        }
        if (studentsTextFilter && !studentsTextFilter.dataset.studentsFilterBound) {
            studentsTextFilter.dataset.studentsFilterBound = '1';
            let studentsTextFilterTimer = null;
            studentsTextFilter.addEventListener('input', () => {
                if (studentsTextFilterTimer) clearTimeout(studentsTextFilterTimer);
                studentsTextFilterTimer = setTimeout(() => {
                    studentsTextFilterTimer = null;
                    renderStudentsTableFromTextarea();
                }, 150);
            });
        }
        if (studentsTable && !studentsTable.dataset.studentsSortBound) {
            studentsTable.dataset.studentsSortBound = '1';
            studentsTable.querySelectorAll('th.is-sortable').forEach((th) => {
                th.addEventListener('click', () => applyStudentsSort(th.dataset.sortKey));
                th.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        applyStudentsSort(th.dataset.sortKey);
                    }
                });
            });
        }
        if (btnAddStudentRow) {
            btnAddStudentRow.addEventListener('click', () => {
                const all = getStudentsFromTextarea();
                all.push({ klasse: '', name: '', email: '' });
                setStudentsTextareaFromRows(all);
                renderStudentsTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnStudentsExportCsv) {
            btnStudentsExportCsv.addEventListener('click', () => {
                const api = window.ms365StudentListImport;
                const rows = getStudentsFromTextarea();
                if (!rows.length) {
                    setSummary('Schülerliste ist leer – nichts zu exportieren.', 'warn');
                    return;
                }
                if (!api || typeof api.exportCsv !== 'function' || !api.exportCsv(rows)) {
                    setSummary('CSV-Export fehlgeschlagen.', 'warn');
                    return;
                }
                setSummary('Schülerliste als CSV exportiert (' + rows.length + ').', 'ok');
            });
        }
        if (btnStudentsExportXlsx) {
            btnStudentsExportXlsx.addEventListener('click', () => {
                const api = window.ms365StudentListImport;
                const rows = getStudentsFromTextarea();
                if (!rows.length) {
                    setSummary('Schülerliste ist leer – nichts zu exportieren.', 'warn');
                    return;
                }
                if (!api || (typeof api.isXlsxReady === 'function' && !api.isXlsxReady())) {
                    setSummary('Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
                    return;
                }
                if (!api || typeof api.exportXlsx !== 'function' || !api.exportXlsx(rows)) {
                    setSummary('XLSX-Export fehlgeschlagen.', 'warn');
                    return;
                }
                setSummary('Schülerliste als XLSX exportiert (' + rows.length + ').', 'ok');
            });
        }
        if (btnVerifyStudentsGraph && !btnVerifyStudentsGraph.dataset.tenantVerifyStudentsBound) {
            btnVerifyStudentsGraph.dataset.tenantVerifyStudentsBound = '1';
            btnVerifyStudentsGraph.addEventListener('click', async () => {
                await runVerifyGraphDirectoryRows(
                    getStudentsFromTextarea(),
                    function (r) {
                        return r.email;
                    },
                    'Schüler:innen',
                    btnVerifyStudentsGraph,
                    function () {
                        renderStudentsTableFromTextarea();
                    }
                );
            });
        }
        if (taStudentCouncil) {
            taStudentCouncil.addEventListener('input', () => {
                studentCouncilSortState.key = null;
                studentCouncilSortState.dir = 1;
                renderStudentCouncilTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnAddStudentCouncilRow) {
            btnAddStudentCouncilRow.addEventListener('click', () => {
                const all = getStudentCouncilFromTextarea();
                all.push({ klasse: '', name: '', email: '' });
                setStudentCouncilTextareaFromRows(all);
                renderStudentCouncilTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnVerifySgaGraph && !btnVerifySgaGraph.dataset.tenantVerifySgaBound) {
            btnVerifySgaGraph.dataset.tenantVerifySgaBound = '1';
            btnVerifySgaGraph.addEventListener('click', async () => {
                await runVerifyGraphDirectoryRows(
                    getSgaFromTextarea(),
                    function (r) { return r.email; },
                    'SGA-Mitglieder',
                    btnVerifySgaGraph,
                    function () { renderSgaTableFromTextarea(); }
                );
            });
        }
        if (btnVerifyStudentCouncilGraph && !btnVerifyStudentCouncilGraph.dataset.tenantVerifyStudentCouncilBound) {
            btnVerifyStudentCouncilGraph.dataset.tenantVerifyStudentCouncilBound = '1';
            btnVerifyStudentCouncilGraph.addEventListener('click', async () => {
                await runVerifyGraphDirectoryRows(
                    getStudentCouncilFromTextarea(),
                    function (r) { return r.email; },
                    'Schülervertretung',
                    btnVerifyStudentCouncilGraph,
                    function () { renderStudentCouncilTableFromTextarea(); }
                );
            });
        }

        if (btnVerifySgaGroup && !btnVerifySgaGroup.dataset.tenantVerifySgaGroupBound) {
            btnVerifySgaGroup.dataset.tenantVerifySgaGroupBound = '1';
            btnVerifySgaGroup.addEventListener('click', async () => {
                if (btnVerifySgaGroup.disabled) return;
                try {
                    btnVerifySgaGroup.disabled = true;
                    btnVerifySgaGroup.setAttribute('aria-busy', 'true');
                    setSummary('SGA-Gruppe: Prüfe in Microsoft 365 …', 'warn');
                    await verifySgaGroupExistence();
                } catch (e) {
                    setSummary('SGA-Gruppe: ' + (e && e.message ? e.message : String(e)), 'warn');
                } finally {
                    btnVerifySgaGroup.disabled = false;
                    btnVerifySgaGroup.removeAttribute('aria-busy');
                }
            });
        }

        if (btnCreateSgaGroup && !btnCreateSgaGroup.dataset.tenantCreateSgaGroupBound) {
            btnCreateSgaGroup.dataset.tenantCreateSgaGroupBound = '1';
            btnCreateSgaGroup.addEventListener('click', async () => {
                if (btnCreateSgaGroup.disabled) return;
                try {
                    btnCreateSgaGroup.disabled = true;
                    btnCreateSgaGroup.setAttribute('aria-busy', 'true');
                    setSummary('SGA-Gruppe: Anlege in Microsoft 365 …', 'warn');
                    await createSgaGroupExistence();
                } catch (e) {
                    setSummary('SGA-Gruppe anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
                } finally {
                    btnCreateSgaGroup.disabled = false;
                    btnCreateSgaGroup.removeAttribute('aria-busy');
                }
            });
        }

        if (btnVerifyStudentCouncilGroup && !btnVerifyStudentCouncilGroup.dataset.tenantVerifyStudentCouncilGroupBound) {
            btnVerifyStudentCouncilGroup.dataset.tenantVerifyStudentCouncilGroupBound = '1';
            btnVerifyStudentCouncilGroup.addEventListener('click', async () => {
                if (btnVerifyStudentCouncilGroup.disabled) return;
                try {
                    btnVerifyStudentCouncilGroup.disabled = true;
                    btnVerifyStudentCouncilGroup.setAttribute('aria-busy', 'true');
                    setSummary('Schülervertretung: Prüfe in Microsoft 365 …', 'warn');
                    await verifyStudentCouncilGroupExistence();
                } catch (e) {
                    setSummary(
                        'Schülervertretung: ' + (e && e.message ? e.message : String(e)),
                        'warn'
                    );
                } finally {
                    btnVerifyStudentCouncilGroup.disabled = false;
                    btnVerifyStudentCouncilGroup.removeAttribute('aria-busy');
                }
            });
        }

        if (btnCreateStudentCouncilGroup && !btnCreateStudentCouncilGroup.dataset.tenantCreateStudentCouncilGroupBound) {
            btnCreateStudentCouncilGroup.dataset.tenantCreateStudentCouncilGroupBound = '1';
            btnCreateStudentCouncilGroup.addEventListener('click', async () => {
                if (btnCreateStudentCouncilGroup.disabled) return;
                try {
                    btnCreateStudentCouncilGroup.disabled = true;
                    btnCreateStudentCouncilGroup.setAttribute('aria-busy', 'true');
                    setSummary('Schülervertretung: Anlege in Microsoft 365 …', 'warn');
                    await createStudentCouncilGroupExistence();
                } catch (e) {
                    setSummary(
                        'Schülervertretung anlegen: ' + (e && e.message ? e.message : String(e)),
                        'warn'
                    );
                } finally {
                    btnCreateStudentCouncilGroup.disabled = false;
                    btnCreateStudentCouncilGroup.removeAttribute('aria-busy');
                }
            });
        }

        if (btnSearchStudentCouncilGroup && !btnSearchStudentCouncilGroup.dataset.tenantSearchSvGroupBound) {
            btnSearchStudentCouncilGroup.dataset.tenantSearchSvGroupBound = '1';
            btnSearchStudentCouncilGroup.addEventListener('click', async () => {
                if (btnSearchStudentCouncilGroup.disabled) return;
                try {
                    btnSearchStudentCouncilGroup.disabled = true;
                    await searchSchoolWideGroups('studentCouncil');
                } catch (e) {
                    setSummary(
                        'Schülervertretung Suche: ' + (e && e.message ? e.message : String(e)),
                        'warn'
                    );
                } finally {
                    btnSearchStudentCouncilGroup.disabled = false;
                }
            });
        }
        if (inpStudentCouncilGroupSearch && !inpStudentCouncilGroupSearch.dataset.tenantSearchSvEnterBound) {
            inpStudentCouncilGroupSearch.dataset.tenantSearchSvEnterBound = '1';
            inpStudentCouncilGroupSearch.addEventListener('keydown', (ev) => {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    if (btnSearchStudentCouncilGroup) btnSearchStudentCouncilGroup.click();
                }
            });
        }
        if (btnUnmatchStudentCouncilGroup && !btnUnmatchStudentCouncilGroup.dataset.tenantUnmatchSvBound) {
            btnUnmatchStudentCouncilGroup.dataset.tenantUnmatchSvBound = '1';
            btnUnmatchStudentCouncilGroup.addEventListener('click', () => {
                unmatchSchoolWideGroup('studentCouncil');
            });
        }
        if (btnOpenStudentCouncilGroup && !btnOpenStudentCouncilGroup.dataset.tenantOpenSvBound) {
            btnOpenStudentCouncilGroup.dataset.tenantOpenSvBound = '1';
            btnOpenStudentCouncilGroup.addEventListener('click', () => {
                openSchoolWideGroupInEntra('studentCouncil');
            });
        }

        if (btnSearchSgaGroup && !btnSearchSgaGroup.dataset.tenantSearchSgaGroupBound) {
            btnSearchSgaGroup.dataset.tenantSearchSgaGroupBound = '1';
            btnSearchSgaGroup.addEventListener('click', async () => {
                if (btnSearchSgaGroup.disabled) return;
                try {
                    btnSearchSgaGroup.disabled = true;
                    await searchSchoolWideGroups('sga');
                } catch (e) {
                    setSummary('SGA Suche: ' + (e && e.message ? e.message : String(e)), 'warn');
                } finally {
                    btnSearchSgaGroup.disabled = false;
                }
            });
        }
        if (inpSgaGroupSearch && !inpSgaGroupSearch.dataset.tenantSearchSgaEnterBound) {
            inpSgaGroupSearch.dataset.tenantSearchSgaEnterBound = '1';
            inpSgaGroupSearch.addEventListener('keydown', (ev) => {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    if (btnSearchSgaGroup) btnSearchSgaGroup.click();
                }
            });
        }
        if (btnUnmatchSgaGroup && !btnUnmatchSgaGroup.dataset.tenantUnmatchSgaBound) {
            btnUnmatchSgaGroup.dataset.tenantUnmatchSgaBound = '1';
            btnUnmatchSgaGroup.addEventListener('click', () => {
                unmatchSchoolWideGroup('sga');
            });
        }
        if (btnOpenSgaGroup && !btnOpenSgaGroup.dataset.tenantOpenSgaBound) {
            btnOpenSgaGroup.dataset.tenantOpenSgaBound = '1';
            btnOpenSgaGroup.addEventListener('click', () => {
                openSchoolWideGroupInEntra('sga');
            });
        }

        const btnLifecycleApply = document.getElementById('tenantStudentLifecycleApply');
        if (btnLifecycleApply) btnLifecycleApply.addEventListener('click', () => applyStudentLifecyclePreview());
        const btnLifecycleDismiss = document.getElementById('tenantStudentLifecycleDismiss');
        if (btnLifecycleDismiss) {
            btnLifecycleDismiss.addEventListener('click', () => {
                const host = document.getElementById('tenantStudentLifecycle');
                if (host) host.hidden = true;
            });
        }

        if (taClasses) {
            taClasses.addEventListener('input', () => {
                classesSortState.key = null;
                classesSortState.dir = 1;
                renderClassesTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnAddClassRow) {
            btnAddClassRow.addEventListener('click', () => {
                const all = getClassesFromTextarea();
                all.push({ code: '', year: '', name: '', headName: '', headEmail: '' });
                setClassesTextareaFromRows(all);
                renderClassesTableFromTextarea();
                scheduleAutoSave();
            });
        }
        if (btnVerifyClassesGraph && !btnVerifyClassesGraph.dataset.tenantVerifyClassesBound) {
            btnVerifyClassesGraph.dataset.tenantVerifyClassesBound = '1';
            btnVerifyClassesGraph.addEventListener('click', async () => {
                const rows = getClassesFromTextarea();
                if (!rows.length) {
                    setSummary('Keine Klassen zum Prüfen vorhanden.', 'warn');
                    return;
                }
                const btn = btnVerifyClassesGraph;
                let found = 0;
                let missed = 0;
                let skipped = 0;
                const seen = new Set();
                try {
                    btn.disabled = true;
                    btn.setAttribute('aria-busy', 'true');
                    setSummary('Klassengruppen-Abgleich läuft …', 'warn');
                    let token = null;
                    for (let i = 0; i < rows.length; i++) {
                        const key = classMatchKey(rows[i]);
                        if (!key) {
                            skipped++;
                            continue;
                        }
                        if (seen.has(key)) continue;
                        seen.add(key);
                        if (!token) {
                            try {
                                token = await graphApi().getGraphToken();
                            } catch {
                                token = null;
                            }
                        }
                        const res = await verifyClassGroupForRow(rows[i], token ? { token: token } : {});
                        if (res && res.skipped) skipped++;
                        else if (res && res.found) found++;
                        else missed++;
                        if (typeof graphApi().sleep === 'function' && i < rows.length - 1) {
                            await graphApi().sleep(300);
                        }
                    }
                    renderClassesTableFromTextarea();
                    renderStatusOverview();
                    setSummary(
                        'Klassen: ' + found + ' gefunden, ' + missed + ' nicht gefunden, ' + skipped + ' ohne Abgleichsschlüssel',
                        missed ? 'warn' : 'ok'
                    );
                } catch (e) {
                    setSummary('Klassengruppen-Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
                } finally {
                    btn.disabled = false;
                    btn.removeAttribute('aria-busy');
                }
            });
        }
        if (btnEnrichClassesFromGroups && !btnEnrichClassesFromGroups.dataset.tenantEnrichClassesBound) {
            btnEnrichClassesFromGroups.dataset.tenantEnrichClassesBound = '1';
            btnEnrichClassesFromGroups.addEventListener('click', async () => {
                const rows = getClassesFromTextarea();
                if (!rows.length) {
                    setSummary('Keine Klassen zum Ergänzen vorhanden.', 'warn');
                    return;
                }
                const btn = btnEnrichClassesFromGroups;
                try {
                    btn.disabled = true;
                    btn.setAttribute('aria-busy', 'true');
                    let working = rows;
                    const local = applyClassEnrichFromLocalLinks(working, { persist: false });
                    working = local.classes || working;
                    const ctx = buildClassEnrichContext();
                    const graphResult = await enrichClassesKvFromGraphOwners(working, {
                        classTeams: ctx.classTeams,
                        teachers: getTeachersFromTextarea(),
                        fetchGroupOwners: graphApi().fetchGroupOwners,
                        getGraphToken: function () {
                            return graphApi().getGraphToken();
                        }
                    });
                    working = graphResult.classes || working;
                    const cur = load();
                    const saved = save(Object.assign({}, cur, { classes: working }));
                    if (taClasses) taClasses.value = classesToLines(saved.classes || []);
                    renderClassesTableFromTextarea();
                    renderStatusOverview();
                    const loc = local.stats || {};
                    const gr = graphResult.stats || {};
                    setSummary(
                        'Ergänzt: ' +
                            (loc.yearFilled || 0) +
                            ' Abschlussjahr(e), ' +
                            ((loc.kvFilled || 0) + (gr.kvFilled || 0)) +
                            ' Klassenvorstand-Daten (lokal + ggf. Graph-Besitzer). Bitte prüfen und speichern ist bereits übernommen.',
                        'ok'
                    );
                    dispatchTenantSettingsChanged(saved, 'class-enrich-groups');
                } catch (e) {
                    setSummary('Ergänzen fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
                } finally {
                    btn.disabled = false;
                    btn.removeAttribute('aria-busy');
                }
            });
        }

        const btnStatusRefresh = document.getElementById('tenantStatusRefresh');
        if (btnStatusRefresh) {
            btnStatusRefresh.addEventListener('click', () => renderStatusOverview());
        }

        window.addEventListener('ms365-tenant-settings-changed', function (ev) {
            const d = ev && ev.detail;
            if (!d) return;
            if (d.reason === 'render' || d.reason === 'autosave') return;
            if (d.source === 'browser-backup-import' || d.source === 'spo-auto-pull') {
                renderFromStorage();
            }
        });

        renderFromStorage();

        // Schritt 4 / 5: Tenant-IST, Match (Dropdown+Speichern), Differenz + Graph-Anlage
        function setTenantDeltaProgress(visible, text, ratio) {
            const wrap = document.getElementById('tenantDeltaProgressWrap');
            const txt = document.getElementById('tenantDeltaProgressText');
            const bar = document.getElementById('tenantDeltaProgressBar');
            const pct = document.getElementById('tenantDeltaProgressPct');
            if (wrap) wrap.style.display = visible ? '' : 'none';
            if (txt && text) txt.textContent = String(text);
            const r = typeof ratio === 'number' && isFinite(ratio) ? Math.max(0, Math.min(1, ratio)) : null;
            if (bar) bar.style.width = r === null ? '0%' : String(Math.round(r * 100)) + '%';
            if (pct) pct.textContent = r === null ? '–' : String(Math.round(r * 100)) + ' %';
        }

        function collectTenantDeltaItems(api) {
            const out = [];
            if (!api || typeof api.loadStructureState !== 'function') return out;
            const { rows } = api.loadStructureState();
            const links = api.loadMatchLinks() || {};
            (rows || []).forEach((r) => {
                if (!r) return;
                const t = String(r.typ || '');
                const gid =
                    normStr(r.tenantGroupId) ||
                    (links[String(r.id)] && normStr(links[String(r.id)].tenantGroupId));
                const groupLike =
                    t === 'Gruppe' ||
                    t === 'Arbeitsgemeinschaft' ||
                    t === 'Klasse' ||
                    t === 'Jahrgang' ||
                    t === 'SchuelerInnen' ||
                    t === 'LehrerInnen' ||
                    t === 'Verwaltung';
                if (groupLike) {
                    if (!gid) {
                        const sug = typeof api.computeCreateSuggestion === 'function' ? api.computeCreateSuggestion(r) : null;
                        const can = !!(sug && sug.displayName && sug.mailNick);
                        let action = can
                            ? 'M365‑Gruppe/Team per Graph anlegen (Mail‑Nickname vorhanden).'
                            : 'Schema/Bezeichnung prüfen (Mail‑Nickname leer).';
                        if (t === 'Klasse' && can) {
                            action = 'Klassen‑Team/Gruppe per Graph anlegen (Schema: meist Microsoft‑365‑Gruppe).';
                        } else if (t === 'Jahrgang' && can) {
                            action = 'Jahrgangs‑Gruppe per Graph anlegen (Schema: Jahrgang/Jg‑Suffix oder Anzeigename).';
                        } else if (t === 'Arbeitsgemeinschaft' && can) {
                            action = 'ARGE als Microsoft‑365‑Gruppe per Graph anlegen.';
                        }
                        out.push({
                            kind: 'group',
                            r,
                            action,
                            canProvision: can
                        });
                    }
                    return;
                }
                if (t === 'Kursteam') {
                    if (gid) return;
                    const kt =
                        window.ms365StructureRules &&
                        typeof window.ms365StructureRules.resolveKursteamKlasseFach === 'function'
                            ? window.ms365StructureRules.resolveKursteamKlasseFach(r, rows)
                            : {
                                  klasse: normStr(r.ktKlasse),
                                  fach: normStr(r.ktFach),
                                  hasBoth: !!(normStr(r.ktKlasse) && normStr(r.ktFach))
                              };
                    if (!kt.hasBoth) {
                        out.push({
                            kind: 'group',
                            r,
                            action:
                                'Kursteam: Klasse und Fach fehlen (Feld „Klasse“/„Fach“ oder Kursteam unter einer „Klasse“-Zeile mit gesetztem Fach).',
                            canProvision: false
                        });
                        return;
                    }
                    const sug = typeof api.computeCreateSuggestion === 'function' ? api.computeCreateSuggestion(r) : null;
                    const can = !!(sug && sug.displayName && sug.mailNick);
                    out.push({
                        kind: 'group',
                        r,
                        action: can ? 'Kursteam (Unified + Team) per Graph anlegen.' : 'Kursteam: Mail‑Nickname/Schema prüfen.',
                        canProvision: can
                    });
                    return;
                }
                const linkRow = links[String(r.id)] || null;
                const linkedUserForPerson = normStr(r.tenantUserId) || (linkRow && normStr(linkRow.tenantUserId));
                if (t === 'Person' && !linkedUserForPerson) {
                    const dn = normStr(r.personName) || normStr(r.bezeichnung);
                    const em = normStr(r.personEmail).toLowerCase();
                    const can = !!(dn && em && em.indexOf('@') !== -1);
                    out.push({
                        kind: 'person',
                        r,
                        action: can
                            ? 'Entra‑Benutzer per Graph anlegen (Name + E‑Mail vorhanden).'
                            : 'Person: Name und gültige E‑Mail/UPN in Schritt 3 ergänzen.',
                        canProvision: can
                    });
                }
            });
            return out;
        }

        function renderTenantInventoryMatchRows() {
            const tbody = document.getElementById('tenantInvMatchTbody');
            const sum = document.getElementById('tenantInvSummary');
            if (!tbody) return;
            const api = window.ms365TenantInventory;
            if (!api || typeof api.loadStructureState !== 'function') {
                tbody.innerHTML =
                    '<tr><td colspan="3" class="muted">Technik wird geladen … Seite ggf. neu laden, falls diese Zeile bleibt.</td></tr>';
                return;
            }
            const { rows } = api.loadStructureState();
            const links = api.loadMatchLinks() || {};
            const cache = typeof api.readCache === 'function' ? api.readCache() : { rows: [], users: [] };
            const groups = cache.rows || [];
            const users = cache.users || [];
            const orgRows = (rows || []).filter((r) => {
                if (!r) return false;
                const t = String(r.typ || '');
                return (
                    t === 'SchuelerInnen' ||
                    t === 'LehrerInnen' ||
                    t === 'Verwaltung' ||
                    t === 'Jahrgang' ||
                    t === 'Klasse' ||
                    t === 'Gruppe' ||
                    t === 'Kursteam' ||
                    t === 'Person' ||
                    t === 'Arbeitsgemeinschaft'
                );
            });
            if (!orgRows.length) {
                tbody.innerHTML =
                    '<tr><td colspan="3" class="muted">Keine Einträge vom Typ Jahrgang, Klasse, Gruppe, Kursteam, ARGE (Arbeitsgemeinschaft) oder Person in der SOLL‑Struktur (Schritt 3).</td></tr>';
                if (sum) {
                    sum.style.display = '';
                    sum.textContent =
                        'Cache: ' + groups.length + ' Gruppe(n)/Team(s), ' + users.length + ' Benutzerkonto/-konten (nach „Tenant laden“).';
                }
                return;
            }
            tbody.replaceChildren();
            orgRows.forEach((r) => {
                const tr = document.createElement('tr');
                const rid = String(r.id || '');
                const linked =
                    normStr(r.tenantGroupId) ||
                    (links[rid] && normStr(links[rid].tenantGroupId)) ||
                    '';
                const c1 = document.createElement('td');
                c1.textContent = (r.bezeichnung || '–') + ' · ' + (r.typ || '');
                const c2 = document.createElement('td');
                const c3 = document.createElement('td');
                c3.className = 'tenant-inv-actions';
                const t = String(r.typ || '');
                if (t === 'Person') {
                    const pNote = document.createElement('div');
                    pNote.className = 'muted';
                    pNote.style.fontSize = '0.82em';
                    pNote.style.lineHeight = '1.35';
                    const linkP = links[rid] || null;
                    const uidShow = normStr(r.tenantUserId) || (linkP && normStr(linkP.tenantUserId));
                    let label = '';
                    if (uidShow) {
                        const hit = users.find(function (u) {
                            return u && String(u.id) === String(uidShow);
                        });
                        label = hit
                            ? (normStr(hit.displayName) || normStr(hit.userPrincipalName) || uidShow) + ' · ' + uidShow
                            : uidShow;
                    }
                    pNote.textContent = uidShow ? 'Entra: ' + label : '— · Benutzer (Abgleich / Schritt 5)';
                    c2.appendChild(pNote);
                    c3.textContent = '';
                } else {
                    const sel = document.createElement('select');
                    sel.setAttribute('data-tenant-inv-sel', rid);
                    const o0 = document.createElement('option');
                    o0.value = '';
                    o0.textContent = '(keine Entra‑Gruppe)';
                    sel.appendChild(o0);
                    groups.forEach((g) => {
                        const o = document.createElement('option');
                        o.value = String(g.id || '');
                        o.textContent = (g.bezeichnung || '(ohne Name)') + ' · ' + (g.typ || '') + (g.alias ? ' · ' + g.alias : '');
                        sel.appendChild(o);
                    });
                    sel.value = linked || '';
                    c2.appendChild(sel);
                    const bSug = document.createElement('button');
                    bSug.type = 'button';
                    bSug.className = 'btn small-btn tenant-inv-icon-btn';
                    bSug.setAttribute('data-tenant-inv-suggest', rid);
                    bSug.setAttribute('aria-label', 'Vorschlag');
                    bSug.title = 'Vorschlag ins Dropdown (ohne Speichern)';
                    bSug.innerHTML = '<i class="bi bi-magic"></i>';
                    const bSave = document.createElement('button');
                    bSave.type = 'button';
                    bSave.className = 'btn btn-success small-btn tenant-inv-icon-btn';
                    bSave.setAttribute('data-tenant-inv-save', rid);
                    bSave.setAttribute('aria-label', 'Speichern');
                    bSave.title = 'Verknüpfung speichern';
                    bSave.innerHTML = '<i class="bi bi-check2"></i>';
                    const bClr = document.createElement('button');
                    bClr.type = 'button';
                    bClr.className = 'btn small-btn tenant-inv-icon-btn';
                    bClr.setAttribute('data-tenant-inv-clear', rid);
                    bClr.setAttribute('aria-label', 'Verknüpfung löschen');
                    bClr.title = 'Verknüpfung löschen';
                    bClr.innerHTML = '<i class="bi bi-x-lg"></i>';
                    c3.appendChild(bSug);
                    c3.appendChild(bSave);
                    c3.appendChild(bClr);
                }
                tr.appendChild(c1);
                tr.appendChild(c2);
                tr.appendChild(c3);
                tbody.appendChild(tr);
            });
            if (sum) {
                sum.style.display = '';
                sum.textContent =
                    'Cache: ' +
                    groups.length +
                    ' Gruppe(n)/Team(s), ' +
                    users.length +
                    ' Benutzer · ' +
                    orgRows.length +
                    ' Match‑Zeilen.';
            }
        }

        function renderTenantDeltaRows() {
            const tbody = document.getElementById('tenantDeltaTbody');
            const sum = document.getElementById('tenantDeltaSummary');
            if (!tbody) return;
            const api = window.ms365TenantInventory;
            if (!api || typeof api.loadStructureState !== 'function') {
                tbody.innerHTML = '<tr><td colspan="3" class="muted">–</td></tr>';
                return;
            }
            const items = collectTenantDeltaItems(api);
            tbody.replaceChildren();
            if (!items.length) {
                tbody.innerHTML =
                    '<tr><td colspan="3" class="muted">Keine offenen Differenzen (Kursteams nur mit gültigem Klasse‑/Fach‑Kontext; Personen nur mit Name und E‑Mail).</td></tr>';
            } else {
                items.forEach((it) => {
                    const tr = document.createElement('tr');
                    const a = document.createElement('td');
                    a.textContent = it.r.bezeichnung || '–';
                    const b = document.createElement('td');
                    b.textContent = it.r.typ || '';
                    const c = document.createElement('td');
                    c.style.display = 'flex';
                    c.style.flexWrap = 'wrap';
                    c.style.gap = '8px';
                    c.style.alignItems = 'center';
                    const span = document.createElement('span');
                    span.textContent = it.action;
                    span.style.flex = '1';
                    span.style.minWidth = '120px';
                    c.appendChild(span);
                    if (it.canProvision && typeof api.provisionGroupRow === 'function' && it.kind === 'group') {
                        const one = document.createElement('button');
                        one.type = 'button';
                        one.className = 'btn btn-success small-btn';
                        one.setAttribute('data-tenant-delta-one-group', String(it.r.id));
                        one.textContent = 'Jetzt anlegen';
                        c.appendChild(one);
                    }
                    if (it.canProvision && typeof api.provisionPersonRow === 'function' && it.kind === 'person') {
                        const one = document.createElement('button');
                        one.type = 'button';
                        one.className = 'btn btn-success small-btn';
                        one.setAttribute('data-tenant-delta-one-person', String(it.r.id));
                        one.textContent = 'Jetzt anlegen';
                        c.appendChild(one);
                    }
                    tr.appendChild(a);
                    tr.appendChild(b);
                    tr.appendChild(c);
                    tbody.appendChild(tr);
                });
            }
            if (sum) {
                sum.style.display = '';
                const prov = items.filter((x) => x.canProvision).length;
                sum.textContent =
                    items.length +
                    ' offene Position(en), davon ' +
                    prov +
                    ' mit Graph‑Anlage möglich (laut Schema).';
            }
        }

        window.__ms365TenantStepsRefreshMatch = function () {
            renderTenantInventoryMatchRows();
            renderTenantDeltaRows();
        };

        const invTbody = document.getElementById('tenantInvMatchTbody');
        if (invTbody && !invTbody.dataset.tenantInvBound) {
            invTbody.dataset.tenantInvBound = '1';
            invTbody.addEventListener('click', (ev) => {
                const api = window.ms365TenantInventory;
                if (!api || typeof api.saveMatchLink !== 'function') return;
                const t = ev.target;
                const saveB = t.closest && t.closest('[data-tenant-inv-save]');
                const sugB = t.closest && t.closest('[data-tenant-inv-suggest]');
                const clrB = t.closest && t.closest('[data-tenant-inv-clear]');
                if (sugB) {
                    const rid = sugB.getAttribute('data-tenant-inv-suggest');
                    const row = (api.loadStructureState().rows || []).find((x) => String(x.id) === String(rid));
                    if (!row) return;
                    const gid = typeof api.suggestGroupForUnit === 'function' ? api.suggestGroupForUnit(row) : '';
                    const sel = invTbody.querySelector('select[data-tenant-inv-sel="' + rid + '"]');
                    if (sel && gid) sel.value = gid;
                    return;
                }
                if (clrB) {
                    const rid = clrB.getAttribute('data-tenant-inv-clear');
                    api.saveMatchLink(rid, '', '');
                    try {
                        if (typeof api.patchStructureRow === 'function') {
                            api.patchStructureRow(rid, { tenantGroupId: '', tenantMailNickname: '', syncStatus: 'Ausstehend' });
                        }
                    } catch {
                        // ignore
                    }
                    window.__ms365TenantStepsRefreshMatch();
                    return;
                }
                if (saveB) {
                    const rid = saveB.getAttribute('data-tenant-inv-save');
                    const sel = invTbody.querySelector('select[data-tenant-inv-sel="' + rid + '"]');
                    const gid = sel && sel.value ? String(sel.value).trim() : '';
                    api.saveMatchLink(rid, gid, '');
                    if (typeof api.patchStructureRow === 'function') {
                        if (gid) {
                            api.patchStructureRow(rid, {
                                tenantGroupId: gid,
                                syncStatus: 'Ok',
                                letzteFehlermeldung: ''
                            });
                        } else {
                            api.patchStructureRow(rid, {
                                tenantGroupId: '',
                                tenantMailNickname: '',
                                syncStatus: 'Ausstehend',
                                letzteFehlermeldung: ''
                            });
                        }
                    }
                    window.__ms365TenantStepsRefreshMatch();
                }
            });
        }

        const deltaTbody = document.getElementById('tenantDeltaTbody');
        if (deltaTbody && !deltaTbody.dataset.tenantDeltaBound) {
            deltaTbody.dataset.tenantDeltaBound = '1';
            deltaTbody.addEventListener('click', async (ev) => {
                const api = window.ms365TenantInventory;
                if (!api) return;
                const gBtn = ev.target.closest && ev.target.closest('[data-tenant-delta-one-group]');
                const pBtn = ev.target.closest && ev.target.closest('[data-tenant-delta-one-person]');
                const id = (gBtn && gBtn.getAttribute('data-tenant-delta-one-group')) || (pBtn && pBtn.getAttribute('data-tenant-delta-one-person'));
                if (!id) return;
                const row = (api.loadStructureState().rows || []).find((x) => String(x.id) === String(id));
                if (!row) return;
                try {
                    setTenantDeltaProgress(true, 'Anlegen …', 0.2);
                    if (gBtn) await api.provisionGroupRow(row);
                    if (pBtn) {
                        const res = await api.provisionPersonRow(row, {});
                        if (res && res.tempPassword) {
                            await dlgAlert('Benutzer angelegt. Einmaliges Kennwort:\n\n' + res.tempPassword, {
                                title: 'Kennwort notieren',
                                okText: 'Verstanden'
                            });
                        }
                    }
                } catch (e) {
                    await dlgAlert('Fehler: ' + (e && e.message ? e.message : String(e)), { title: 'Fehler' });
                } finally {
                    setTenantDeltaProgress(false, '', null);
                    window.__ms365TenantStepsRefreshMatch();
                }
            });
        }

        async function runTenantInventoryAutomatch() {
            const api = window.ms365TenantInventory;
            const elSt = document.getElementById('tenantInvStatus');
            if (!api || typeof api.loadStructureState !== 'function' || typeof api.suggestGroupForUnit !== 'function') {
                if (elSt) elSt.textContent = 'Automatching: Technik nicht bereit.';
                return;
            }
            const cache = typeof api.readCache === 'function' ? api.readCache() : { rows: [] };
            const gr = cache.rows || [];
            if (!gr.length) {
                await dlgAlert('Zuerst „Tenant laden“ ausführen, damit Entra-Gruppen für das Automatching vorliegen.', {
                    title: 'Automatching'
                });
                return;
            }
            const gidSet = new Set(gr.map((g) => String(g.id || '').trim()).filter(Boolean));
            const st0 = api.loadStructureState();
            const rows = st0.rows || [];
            let links = api.loadMatchLinks() || {};
            const orgRows = rows.filter((r) => {
                if (!r || r.isStructureTreeRoot) return false;
                const ty = String(r.typ || '');
                return (
                    ty === 'Jahrgang' ||
                    ty === 'Klasse' ||
                    ty === 'Gruppe' ||
                    ty === 'Kursteam' ||
                    ty === 'Arbeitsgemeinschaft'
                );
            });
            let saved = 0;
            let skippedHas = 0;
            let skippedNoHit = 0;
            for (let i = 0; i < orgRows.length; i++) {
                const r = orgRows[i];
                const rid = String(r.id || '');
                const linked =
                    normStr(r.tenantGroupId) ||
                    (links[rid] && normStr(links[rid].tenantGroupId));
                if (linked) {
                    skippedHas++;
                    continue;
                }
                const gid = String(api.suggestGroupForUnit(r) || '').trim();
                if (!gid || !gidSet.has(gid)) {
                    skippedNoHit++;
                    continue;
                }
                try {
                    api.saveMatchLink(rid, gid, 'Automatching');
                    if (typeof api.patchStructureRow === 'function') {
                        api.patchStructureRow(rid, {
                            tenantGroupId: gid,
                            syncStatus: 'Ok',
                            letzteFehlermeldung: ''
                        });
                    }
                    saved++;
                    links = api.loadMatchLinks() || links;
                    r.tenantGroupId = gid;
                } catch {
                    skippedNoHit++;
                }
            }
            window.__ms365TenantStepsRefreshMatch();
            const parts = ['Automatching: ' + saved + ' neu verknüpft.'];
            if (skippedHas) parts.push(skippedHas + ' bereits gesetzt.');
            if (skippedNoHit) parts.push(skippedNoHit + ' ohne Treffer.');
            if (elSt) elSt.textContent = parts.join(' ');
        }

        const invAuto = document.getElementById('tenantInvAutoMatchBtn');
        if (invAuto && !invAuto.dataset.tenantInvAutoBound) {
            invAuto.dataset.tenantInvAutoBound = '1';
            invAuto.addEventListener('click', () => void runTenantInventoryAutomatch());
        }

        const invBtn = document.getElementById('tenantInvRefreshBtn');
        if (invBtn) {
            invBtn.addEventListener('click', async () => {
                const st = document.getElementById('tenantInvStatus');
                const api = window.ms365TenantInventory;
                if (!api || typeof api.refresh !== 'function') {
                    if (st) st.textContent = 'Schulstruktur‑Modul nicht geladen. Seite neu laden.';
                    return;
                }
                invBtn.disabled = true;
                if (st) st.textContent = 'Lade Daten über Microsoft Graph …';
                try {
                    await api.refresh((ev) => {
                        if (!st) return;
                        const ph = ev && ev.phase === 'users' ? 'Benutzer' : 'Gruppen/Teams';
                        const pg = ev && ev.page != null ? ' (Seite ' + ev.page + ')' : '';
                        st.textContent = ph + ' werden geladen' + pg + ' …';
                    });
                    const c = api.readCache();
                    if (st) {
                        st.textContent =
                            'Fertig: ' +
                            (c.rows || []).length +
                            ' Gruppe(n)/Team(s), ' +
                            (c.users || []).length +
                            ' Benutzerkonto/-konten.';
                    }
                    renderTenantInventoryMatchRows();
                    renderTenantDeltaRows();
                } catch (e) {
                    if (st) st.textContent = 'Fehler: ' + (e && e.message ? e.message : String(e));
                } finally {
                    invBtn.disabled = false;
                }
            });
        }
        const deltaBtn = document.getElementById('tenantDeltaRefreshBtn');
        if (deltaBtn) {
            deltaBtn.addEventListener('click', () => renderTenantDeltaRows());
        }

        async function runBatchProvision(kind) {
            const api = window.ms365TenantInventory;
            if (!api) return;
            const items = collectTenantDeltaItems(api).filter((x) => x.canProvision && x.kind === kind);
            if (!items.length) {
                await dlgAlert(
                    kind === 'group'
                        ? 'Keine anlegbaren Gruppen (Jahrgang, Klasse, Gruppe, Kursteam, ARGE – Schema prüfen).'
                        : 'Keine anlegbaren Personen (Name + E‑Mail).',
                    { title: 'Delta-Anlage' }
                );
                return;
            }
            const n = items.length;
            const ok = await dlgConfirm(
                kind === 'group'
                    ? n + ' Gruppe(n)/Team(s) (inkl. Klasse, Jahrgang, ARGE, Kursteam) wirklich in Entra anlegen?'
                    : n + ' Benutzerkonto/-konten wirklich in Entra anlegen?',
                { title: 'Entra-Anlage', okText: 'Anlegen', danger: true }
            );
            if (!ok) return;
            const btnG = document.getElementById('tenantDeltaProvisionGroupsBtn');
            const btnP = document.getElementById('tenantDeltaProvisionPersonsBtn');
            if (btnG) btnG.disabled = true;
            if (btnP) btnP.disabled = true;
            let pwdNotes = [];
            for (let i = 0; i < items.length; i++) {
                const it = items[i];
                const ratio = (i + 0.35) / n;
                setTenantDeltaProgress(true, 'Anlegen ' + (i + 1) + ' / ' + n + ' …', ratio);
                try {
                    if (kind === 'group') await api.provisionGroupRow(it.r);
                    else {
                        const res = await api.provisionPersonRow(it.r, { skipConfirm: true });
                        if (res && res.tempPassword) pwdNotes.push((it.r.bezeichnung || it.r.personName || '') + ': ' + res.tempPassword);
                    }
                } catch (e) {
                    await dlgAlert('Abbruch bei Position ' + (i + 1) + ': ' + (e && e.message ? e.message : String(e)), {
                        title: 'Fehler'
                    });
                    break;
                }
            }
            if (btnG) btnG.disabled = false;
            if (btnP) btnP.disabled = false;
            setTenantDeltaProgress(true, 'Fertig.', 1);
            setTimeout(() => setTenantDeltaProgress(false, '', null), 1400);
            if (pwdNotes.length) {
                await dlgAlert('Temporäre Kennwörter (bitte sicher notieren):\n\n' + pwdNotes.join('\n'), {
                    title: 'Kennwörter',
                    okText: 'Verstanden'
                });
            }
            window.__ms365TenantStepsRefreshMatch();
        }

        const btnBatchG = document.getElementById('tenantDeltaProvisionGroupsBtn');
        if (btnBatchG) btnBatchG.addEventListener('click', () => runBatchProvision('group'));
        const btnBatchP = document.getElementById('tenantDeltaProvisionPersonsBtn');
        if (btnBatchP) btnBatchP.addEventListener('click', () => runBatchProvision('person'));

        renderTenantInventoryMatchRows();
        renderTenantDeltaRows();

        function webuntisMergeDeps() {
            return {
                webuntis: window.ms365WebuntisExportImport,
                schoolSis: window.ms365SchoolSisImport,
                parseTeachersLines:
                    typeof window.ms365TenantSettingsParseTeachersLines === 'function'
                        ? window.ms365TenantSettingsParseTeachersLines
                        : function () {
                              return [];
                          },
                parseStudentsLines:
                    typeof window.ms365TenantSettingsParseStudentsLines === 'function'
                        ? window.ms365TenantSettingsParseStudentsLines
                        : function () {
                              return [];
                          },
                parseSubjectsLines:
                    typeof window.ms365TenantSettingsParseSubjectsLines === 'function'
                        ? window.ms365TenantSettingsParseSubjectsLines
                        : function () {
                              return [];
                          },
                parseClassesLines:
                    typeof window.ms365TenantSettingsParseClassesLines === 'function'
                        ? window.ms365TenantSettingsParseClassesLines
                        : function () {
                              return [];
                          }
            };
        }

        async function applyWebuntisImportPayload(payload) {
            if (!payload) return;
            const merged = mergeWebuntisImportWithExisting(
                {
                    teachersLines: taTeachers ? taTeachers.value : '',
                    studentsLines: taStudents ? taStudents.value : '',
                    subjectsLines: taSubjects ? taSubjects.value : '',
                    classesLines: taClasses ? taClasses.value : ''
                },
                payload,
                webuntisMergeDeps()
            );
            const lines = merged.lines || {};
            if (lines.subjectsLines != null && typeof parseLinesToSubjects === 'function') {
                const nextSubjects = parseLinesToSubjects(lines.subjectsLines);
                const ok = await confirmSubjectCatalogIfShrinks(nextSubjects, 'WebUntis-Import');
                if (!ok) {
                    if (summary) {
                        summary.textContent =
                            'WebUntis-Import abgebrochen – die Fächerliste würde Kürzel verlieren. Lehrer/Schüler wurden nicht übernommen.';
                        summary.className = 'summary warn';
                    }
                    return;
                }
            }
            const prevStudents = getStudentsFromTextarea();
            if (taTeachers && lines.teachersLines != null) {
                taTeachers.value = lines.teachersLines;
                teachersSortState.key = null;
                teachersSortState.dir = 1;
                renderTeachersTableFromTextarea();
            }
            if (taStudents && lines.studentsLines != null) {
                taStudents.value = lines.studentsLines;
                studentsSortState.key = null;
                studentsSortState.dir = 1;
                renderStudentsTableFromTextarea();
            }
            if (taSubjects && lines.subjectsLines != null) {
                taSubjects.value = lines.subjectsLines;
                renderSubjectsTableFromTextarea();
            }
            if (taClasses && lines.classesLines != null) {
                taClasses.value = lines.classesLines;
                renderClassesTableFromTextarea();
            }
            const sis = sisApi();
            if (merged.studentDiff && sis) {
                rememberSisHistory(merged.studentDiff, 'webuntis', 'merge');
                renderStudentLifecyclePanel(prevStudents, getStudentsFromTextarea());
            }
            if (summary) {
                const detail = summarizeWebuntisMergeResult(merged.stats, merged.studentDiff, sis || {});
                summary.textContent =
                    'WebUntis-Import mit bestehenden Stammdaten abgeglichen (kein Doppelimport). ' +
                    detail +
                    ' – bitte unten „Speichern“ klicken.';
                summary.className =
                    merged.studentDiff && merged.studentDiff.conflicts && merged.studentDiff.conflicts.length
                        ? 'summary warn'
                        : 'summary ok';
            }
            if (merged.studentDiff && merged.studentDiff.conflicts && merged.studentDiff.conflicts.length) {
                const msgs = merged.studentDiff.conflicts
                    .slice(0, 5)
                    .map(function (c) {
                        return c.summary || c.type;
                    })
                    .join('\n');
                dlgAlert(
                    'Beim Abgleich wurden Konflikte gefunden (bitte in der Schülerliste prüfen):\n\n' + msgs,
                    { title: 'WebUntis-Abgleich', okText: 'Verstanden' }
                );
            }
            scheduleAutoSave();
        }

        const pendingWebuntis = consumeWebuntisImportPayload();
        if (pendingWebuntis) void applyWebuntisImportPayload(pendingWebuntis);

        // Accordion: immer nur EIN Schritt offen (details.step)
        try {
            const steps = Array.from(document.querySelectorAll('details.step'));
            steps.forEach((d) => {
                d.addEventListener('toggle', () => {
                    if (!d.open) return;
                    steps.forEach((o) => {
                        if (o !== d) o.open = false;
                    });
                });
            });
        } catch {
            // ignore
        }
    }

    bindUi();
})();

