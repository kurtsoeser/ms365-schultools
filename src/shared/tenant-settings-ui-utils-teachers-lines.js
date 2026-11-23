/**
 * Teacher list / subject-catalog helpers for tenant-settings-ui.
 */
import {
    subjectCatalogShrinkReport,
    formatSubjectCatalogShrinkConfirmDe,
    subjectsFromImportPayload
} from './subject-list-enrich.js';

import {
    dlgConfirm,
    getField,
    normCode,
    normStr
} from './tenant-settings-ui-utils-dialog.js';

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

export {
    exposeSubjectCatalogGuardOnWindow,
    confirmSubjectCatalogIfShrinks,
    teacherJsonRowsToSemicolonLines,
    teachersRowsToAoa,
    csvEscapeField,
    downloadTeachersCsv,
    compareTeachersDe,
    sortTeacherRows
};
