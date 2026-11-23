/**
 * Student spreadsheet / lines helpers for tenant-settings-ui.
 */
import { getField, normCode, normStr } from './tenant-settings-ui-utils-dialog.js';
import {
    ensureXlsxReady,
    workbookFirstSheetAoa,
    csvTextToWorkbook
} from './tenant-settings-ui-utils-spreadsheet.js';
import { csvEscapeField, compareTeachersDe } from './tenant-settings-ui-utils-teachers-lines.js';
import { schoolDomainForEmail } from './tenant-settings-ui-utils-email-pattern.js';

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
                result.meta.emailHint =
                    'Keine Schuldomain gesetzt: private Kontaktmails (Gmail/GMX/…) werden verworfen, Schul-Mails bleiben. Domain setzen, um fehlende Mails zu erzeugen.';
            } else if (result.meta && result.meta.privateStudentEmailsIgnored > 0) {
                result.meta.emailHint =
                    result.meta.privateStudentEmailsIgnored +
                    ' private Schüler-Kontaktmail(s) verworfen (nicht @' +
                    domain +
                    ').';
            }
            if (result.meta && result.meta.emptyKlasse > 0) {
                result.meta.classHint =
                    result.meta.emptyKlasse +
                    ' ohne Klasse in WebUntis – lokale Klasse bleibt beim Zusammenführen erhalten.';
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

export {
    readFilesToAoaSheets,
    importStudentSheetsResult,
    studentJsonRowsToSemicolonLines,
    studentsRowsToAoa,
    downloadStudentsCsv,
    sortStudentRows
};
