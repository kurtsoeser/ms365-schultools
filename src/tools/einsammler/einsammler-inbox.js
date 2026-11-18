/**
 * Lokaler Posteingang für Eltern-Antworten (gleicher Browser) + Slip-Import.
 */

import { decodeResponseSlip } from './einsammler-token.js';
import { setAntwortStatus, loadState } from './einsammler-store.js';

const INBOX_KEY = 'ms365-einsammler-inbox-v1';

function readInbox() {
    try {
        const raw = localStorage.getItem(INBOX_KEY);
        const arr = raw ? JSON.parse(raw) : [];
        return Array.isArray(arr) ? arr : [];
    } catch {
        return [];
    }
}

function writeInbox(arr) {
    localStorage.setItem(INBOX_KEY, JSON.stringify(arr || []));
}

/**
 * @param {object} slip
 */
export function pushInboxSlip(slip) {
    const list = readInbox();
    list.push({ ...slip, receivedAt: new Date().toISOString() });
    writeInbox(list);
    return list.length;
}

export function listInbox() {
    return readInbox();
}

export function clearInbox() {
    writeInbox([]);
}

/**
 * Übernimmt alle Inbox-Einträge in den Store.
 * @returns {{ applied: number, skipped: number }}
 */
export function mergeInboxIntoStore() {
    const list = readInbox();
    let applied = 0;
    let skipped = 0;
    const keep = [];
    for (const slip of list) {
        const r = applySlip(slip);
        if (r.ok) applied++;
        else {
            skipped++;
            keep.push(slip);
        }
    }
    writeInbox(keep);
    return { applied, skipped };
}

/**
 * @param {object} slip
 * @returns {{ ok: boolean, error?: string }}
 */
export function applySlip(slip) {
    if (!slip || !slip.aid) return { ok: false, error: 'Slip unvollständig.' };
    const state = loadState();
    const ant = (state.antworten || []).find((a) => a.id === slip.aid);
    if (!ant) return { ok: false, error: 'Antwort-ID unbekannt (dieser Browser).' };
    if (slip.eid && ant.einsammlerId !== slip.eid) {
        return { ok: false, error: 'Einsammler-ID passt nicht.' };
    }
    const r = setAntwortStatus(slip.aid, slip.status, slip.kommentar || '', {
        anhangName: slip.anhangName || '',
        anhangMime: slip.anhangMime || '',
        anhangSize: slip.anhangSize || 0,
        anhangKey: slip.anhangKey || slip.aid || ''
    });
    return r.ok ? { ok: true } : { ok: false, error: (r.errors && r.errors[0]) || 'Fehler' };
}

/**
 * @param {string} rawCode
 */
export function importResponseCode(rawCode) {
    const decoded = decodeResponseSlip(rawCode);
    if (!decoded.ok) return decoded;
    const applied = applySlip(decoded.slip);
    if (!applied.ok) return applied;
    return { ok: true, slip: decoded.slip };
}
