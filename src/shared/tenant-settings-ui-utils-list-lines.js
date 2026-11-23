/**
 * Admin / subjects / ARGEs / classes JSON→lines helpers.
 */
import { getField, normCode, normStr } from './tenant-settings-ui-utils-dialog.js';
import { subjectsFromSpreadsheetJsonRows, subjectsToLines as subjectsToLinesPure } from './tenant-settings-ui-subjects.js';

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
    return subjectsToLinesPure(subjectsFromSpreadsheetJsonRows(jsonRows));
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

export {
    adminJsonRowsToSemicolonLines,
    subjectsJsonRowsToSemicolonLines,
    argesJsonRowsToSemicolonLines,
    classesJsonRowsToSemicolonLines
};
