/**
 * Unterrichtsteams-Katalog: Filtern, Sortieren, Zeilen bearbeiten (Unterrichtsbelegung).
 */
import { normalizeBelegungSnapshot, buildSnapshotFromTeamsData } from '../../shared/unterrichtsbelegung-logic.js';
import {
    collectPlannedRowsFromKursteamState,
    readKursteamStateFromBrowserStorage,
    teachingSlotKey
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
            r.lehrerName,
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
    const out = {};
    if (Object.prototype.hasOwnProperty.call(p, 'klasse')) out.klasse = normStr(p.klasse);
    if (Object.prototype.hasOwnProperty.call(p, 'fach')) out.fach = normStr(p.fach);
    if (Object.prototype.hasOwnProperty.call(p, 'lehrerCode')) out.lehrerCode = normCode(p.lehrerCode);
    if (Object.prototype.hasOwnProperty.call(p, 'lehrerName')) out.lehrerName = normStr(p.lehrerName);
    if (Object.prototype.hasOwnProperty.call(p, 'lehrerEmail')) out.lehrerEmail = normMail(p.lehrerEmail);
    if (Object.prototype.hasOwnProperty.call(p, 'gruppe')) out.gruppe = normStr(p.gruppe);
    if (Object.prototype.hasOwnProperty.call(p, 'teamName')) out.teamName = normStr(p.teamName);
    if (Object.prototype.hasOwnProperty.call(p, 'gruppenmail')) out.gruppenmail = normMail(p.gruppenmail);
    if (Object.prototype.hasOwnProperty.call(p, 'graphGroupId')) out.graphGroupId = normStr(p.graphGroupId);
    if (Object.prototype.hasOwnProperty.call(p, 'linkedAt')) out.linkedAt = normStr(p.linkedAt);
    return out;
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

/**
 * Wenn App-Belegung leer ist: Kursteam-Zwischenstand (localStorage) in Snapshot mergen.
 * @param {object|null} existingSnapshot
 * @param {object|null} kursteamState
 */
export function hydrateUnterrichtsbelegungFromKursteamState(existingSnapshot, kursteamState) {
    const { snapshot, imported, source } = mergeWizardStateIntoSnapshot(existingSnapshot, kursteamState);
    if (!imported || !snapshot || !Array.isArray(snapshot.rows) || !snapshot.rows.length) {
        return { ok: false, snapshot: existingSnapshot, imported: 0, source: 'none' };
    }
    return { ok: true, snapshot, imported: snapshot.rows.length, source };
}

/**
 * @param {object|null} container ms365AppDataV2.getContainer()
 * @param {string} [currentYear]
 * @returns {{ year: string, count: number }|null}
 */
export function findUnterrichtsbelegungInOtherYears(container, currentYear) {
    const c = container && typeof container === 'object' ? container : null;
    const by = c && c.years && c.years.byLabel && typeof c.years.byLabel === 'object' ? c.years.byLabel : null;
    if (!by) return null;
    const cur = normStr(currentYear || (c.years && c.years.current));
    let best = null;
    Object.keys(by).forEach((y) => {
        if (normStr(y) === cur) return;
        const snap = by[y] && by[y].unterrichtsbelegung;
        const count = Array.isArray(snap && snap.rows) ? snap.rows.length : 0;
        if (count > 0 && (!best || count > best.count)) {
            best = { year: String(y), count };
        }
    });
    return best;
}

export function countLinked(rows) {
    return (Array.isArray(rows) ? rows : []).filter((r) => normStr(r && r.graphGroupId)).length;
}

/**
 * Graph-IDs aus Abgleich (Mail-Nickname oder Unterrichtsslot) in Katalog-Zeilen schreiben.
 * @param {object[]} rows
 * @param {object[]} matched Ausgabe von buildKursteamAbgleichReport().matched
 * @returns {{ rows: object[], linked: number }}
 */
export function applyAbgleichMatchesToRows(rows, matched) {
    const list = (Array.isArray(rows) ? rows : []).map((r) => Object.assign({}, r));
    const byMail = new Map();
    const bySlot = new Map();
    (Array.isArray(matched) ? matched : []).forEach((m) => {
        if (!m || !normStr(m.graphGroupId)) return;
        const mail = normMail(m.gruppenmail);
        if (mail) byMail.set(mail, m);
        const sk = teachingSlotKey(m);
        if (sk && sk !== '|||') bySlot.set(sk, m);
    });
    let linked = 0;
    list.forEach((r, i) => {
        if (normStr(r.graphGroupId)) return;
        const mail = normMail(r.gruppenmail);
        let hit = mail && byMail.has(mail) ? byMail.get(mail) : null;
        if (!hit) {
            const sk = teachingSlotKey(r);
            hit = sk && sk !== '|||' && bySlot.has(sk) ? bySlot.get(sk) : null;
        }
        if (!hit || !normStr(hit.graphGroupId)) return;
        list[i] = Object.assign({}, r, {
            graphGroupId: normStr(hit.graphGroupId),
            gruppenmail: r.gruppenmail || hit.gruppenmail || '',
            teamName: r.teamName || hit.teamName || ''
        });
        linked += 1;
    });
    return { rows: list, linked };
}
