/**
 * Unterrichtsteams-Katalog: Filtern, Sortieren, Zeilen bearbeiten (Unterrichtsbelegung).
 */
import { normalizeBelegungSnapshot, buildSnapshotFromTeamsData } from '../../shared/unterrichtsbelegung-logic.js';
import {
    collectPlannedRowsFromKursteamState,
    readKursteamStateFromBrowserStorage
} from '../../shared/kursteam-belegung-abgleich-logic.js';
import { mergeBelegungWithGraphImport } from '../../shared/kursteam-graph-import-logic.js';

function normStr(v) {
    return String(v ?? '').trim();
}

function normCode(v) {
    return normStr(v).toUpperCase();
}

function normMail(v) {
    return normStr(v).toLowerCase();
}

function normClass(v) {
    return normStr(v).toUpperCase();
}

/**
 * @param {object} row
 * @returns {string}
 */
export function rowStableKey(row) {
    const r = row && typeof row === 'object' ? row : {};
    const mail = normMail(r.gruppenmail);
    if (mail) return 'm:' + mail;
    const gid = normStr(r.graphGroupId);
    if (gid) return 'g:' + gid;
    return (
        's:' +
        [normClass(r.klasse), normCode(r.lehrerCode), normStr(r.fach).toUpperCase(), normStr(r.gruppe).toUpperCase()].join(
            '|'
        )
    );
}

/**
 * @param {object[]} rows
 */
export function sortUnterrichtsteamRows(rows) {
    return (Array.isArray(rows) ? rows : [])
        .map((r) => Object.assign({}, r))
        .sort((a, b) => {
            const c = normClass(a.klasse).localeCompare(normClass(b.klasse), 'de');
            if (c !== 0) return c;
            const f = normStr(a.fach).localeCompare(normStr(b.fach), 'de');
            if (f !== 0) return f;
            const t = normCode(a.lehrerCode).localeCompare(normCode(b.lehrerCode), 'de');
            if (t !== 0) return t;
            return normMail(a.gruppenmail).localeCompare(normMail(b.gruppenmail), 'de');
        });
}

/**
 * @param {object[]} rows
 * @param {'klasse'|'fach'|'lehrerCode'} field
 */
export function uniqueFieldValues(rows, field) {
    const set = new Set();
    (Array.isArray(rows) ? rows : []).forEach((r) => {
        const v = field === 'lehrerCode' ? normCode(r && r.lehrerCode) : normStr(r && r[field]);
        if (v) set.add(v);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
}

/**
 * @param {object[]} rows
 * @param {{ q?: string, klasse?: string, fach?: string, lehrerCode?: string, linkedOnly?: boolean, unlinkedOnly?: boolean }} filters
 */
export function filterUnterrichtsteamRows(rows, filters) {
    const f = filters && typeof filters === 'object' ? filters : {};
    const q = normStr(f.q).toLowerCase();
    const klasse = normClass(f.klasse);
    const fach = normStr(f.fach);
    const lehrer = normCode(f.lehrerCode);

    return (Array.isArray(rows) ? rows : []).filter((r) => {
        if (klasse && normClass(r.klasse) !== klasse) return false;
        if (fach && normStr(r.fach) !== fach) return false;
        if (lehrer && normCode(r.lehrerCode) !== lehrer) return false;
        const linked = !!normStr(r.graphGroupId);
        if (f.linkedOnly && !linked) return false;
        if (f.unlinkedOnly && linked) return false;
        if (!q) return true;
        const hay = [
            r.klasse,
            r.fach,
            r.lehrerCode,
            r.lehrerEmail,
            r.teamName,
            r.gruppenmail,
            r.graphGroupId
        ]
            .map((x) => String(x || '').toLowerCase())
            .join(' ');
        return hay.indexOf(q) !== -1;
    });
}

/**
 * @param {object} patch
 */
export function normalizeEditableRow(patch) {
    const p = patch && typeof patch === 'object' ? patch : {};
    return {
        klasse: normStr(p.klasse),
        fach: normStr(p.fach),
        lehrerCode: normCode(p.lehrerCode),
        lehrerEmail: normMail(p.lehrerEmail),
        gruppe: normStr(p.gruppe),
        teamName: normStr(p.teamName),
        gruppenmail: normMail(p.gruppenmail),
        graphGroupId: normStr(p.graphGroupId),
        linkedAt: normStr(p.linkedAt)
    };
}

/**
 * @param {object[]} rows
 * @param {string} key
 * @param {object} patch
 */
export function updateRowByKey(rows, key, patch) {
    const k = String(key || '');
    const next = (Array.isArray(rows) ? rows : []).map((r) => Object.assign({}, r));
    const idx = next.findIndex((r) => rowStableKey(r) === k);
    if (idx < 0) return { rows: next, ok: false };
    const merged = Object.assign({}, next[idx], normalizeEditableRow(patch));
    if (merged.gruppenmail) merged.gruppenmail = normMail(merged.gruppenmail);
    next[idx] = merged;
    return { rows: next, ok: true };
}

/**
 * @param {object[]} rows
 * @param {string} key
 */
export function removeRowByKey(rows, key) {
    const k = String(key || '');
    const next = (Array.isArray(rows) ? rows : []).filter((r) => rowStableKey(r) !== k);
    return { rows: next, removed: next.length < (rows || []).length };
}

/**
 * @param {object[]} rows
 * @param {object} [meta]
 */
export function snapshotFromRows(rows, meta) {
    const m = meta && typeof meta === 'object' ? meta : {};
    return normalizeBelegungSnapshot({
        updatedAt: new Date().toISOString(),
        yearPrefix: normStr(m.yearPrefix),
        source: normStr(m.source) || 'unterrichtsteams-katalog',
        rows: sortUnterrichtsteamRows(rows)
    });
}

/**
 * @param {object|null} existingSnapshot
 */
export function mergeWizardStateIntoSnapshot(existingSnapshot, kursteamState) {
    const st = kursteamState && typeof kursteamState === 'object' ? kursteamState : null;
    if (!st) return { snapshot: existingSnapshot, imported: 0, source: 'none' };

    if (st.teamsGenerated && Array.isArray(st.teamsData) && st.teamsData.length) {
        const built = buildSnapshotFromTeamsData(st.teamsData, {
            yearPrefix: normStr(st.yearPrefix),
            source: 'kursteams'
        });
        if (!built || !built.rows.length) return { snapshot: existingSnapshot, imported: 0, source: 'teamsData-empty' };
        const { snapshot } = mergeBelegungWithGraphImport(existingSnapshot, built.rows, {
            yearPrefix: built.yearPrefix,
            source: 'kursteams+unterrichtsteams-katalog'
        });
        return { snapshot, imported: built.rows.length, source: 'teamsData' };
    }

    const planned = collectPlannedRowsFromKursteamState(st);
    if (!planned.rows.length) return { snapshot: existingSnapshot, imported: 0, source: 'none' };
    const { snapshot } = mergeBelegungWithGraphImport(existingSnapshot, planned.rows, {
        yearPrefix: normStr(st.yearPrefix),
        source: 'webuntis+unterrichtsteams-katalog'
    });
    return { snapshot, imported: planned.rows.length, source: planned.source };
}

export function loadWizardMergeFromBrowser() {
    const state = readKursteamStateFromBrowserStorage();
    return state;
}

export function countLinked(rows) {
    return (Array.isArray(rows) ? rows : []).filter((r) => normStr(r && r.graphGroupId)).length;
}
