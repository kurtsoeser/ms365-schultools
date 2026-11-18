/**
 * Lokaler Posteingang für Lesebestätigungen (gleicher Browser).
 */

import { decodeLeseSlip } from './elternbrief-token.js';
import { markGelesen } from './elternbrief-store.js';

const KEY = 'ms365-elternbrief-inbox-v1';

function read() {
    try {
        const raw = localStorage.getItem(KEY);
        const arr = raw ? JSON.parse(raw) : [];
        return Array.isArray(arr) ? arr : [];
    } catch {
        return [];
    }
}

function write(arr) {
    localStorage.setItem(KEY, JSON.stringify(arr || []));
}

/**
 * @param {object} slip
 */
export function pushLeseSlip(slip) {
    const list = read();
    list.push({ ...slip, receivedAt: new Date().toISOString() });
    write(list);
}

export function listInbox() {
    return read();
}

export function mergeInboxIntoStore() {
    const list = read();
    let applied = 0;
    let skipped = 0;
    const keep = [];
    /** @type {object[]} */
    const items = [];
    for (const slip of list) {
        const r = markGelesen(slip.bid, slip.eid);
        if (r.ok) {
            applied++;
            if (r.item) items.push(r.item);
        } else {
            skipped++;
            keep.push(slip);
        }
    }
    write(keep);
    return { applied, skipped, items };
}

/**
 * @param {string} raw
 */
export function importLeseCode(raw) {
    const decoded = decodeLeseSlip(raw);
    if (!decoded.ok) return decoded;
    const r = markGelesen(decoded.slip.bid, decoded.slip.eid);
    return r.ok
        ? { ok: true, item: r.item || null }
        : { ok: false, error: (r.errors && r.errors[0]) || 'Fehler', item: null };
}
