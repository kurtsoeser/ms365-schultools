/**
 * Reine Helfer für Klassengruppen / Jahrgangsgruppen (Analyse 02 Phase B).
 * Move-first aus jahrgangsgruppen.js – Signaturen an Closure angepasst (explizite Params).
 */
import { normStr, normCode, normEmail } from '../../shared/utils/strings.js';

export { normStr, normCode, normEmail };

export function sanitizeNick(raw) {
    return String(raw || '')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9-]/g, '')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '')
        .slice(0, 60);
}

export function rowKey(row) {
    return normCode(row && row.code) || normStr(row && row.name).toUpperCase();
}

export function domainFromMail(mail) {
    const m = String(mail || '')
        .trim()
        .toLowerCase();
    const i = m.lastIndexOf('@');
    return i >= 0 ? m.slice(i + 1) : '';
}

export function isDirektionRole(roleRaw) {
    const r = normStr(roleRaw).toLowerCase();
    return !!r && (r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1);
}

export function classCodeExists(classes, code, exceptCode) {
    const key = normCode(code);
    const skip = normCode(exceptCode);
    return (classes || []).some(function (r) {
        const c = normCode(r.code);
        if (skip && c === skip) return false;
        return c === key;
    });
}

export function remapStudentKlassen(list, fromCode, toCode) {
    const from = normCode(fromCode);
    const to = normCode(toCode);
    if (!from || from === to) return Array.isArray(list) ? list.slice() : [];
    return (Array.isArray(list) ? list : []).map(function (s) {
        if (normCode(s && s.klasse) !== from) return s;
        return Object.assign({}, s, { klasse: to });
    });
}

/**
 * Ableitung ohne window-Hook; Caller kann stable-Nickname-API vorher anwenden.
 */
export function deriveNickFallback(row) {
    if (!row) return '';
    const fromRow = sanitizeNick(row.stableMailNickname);
    if (fromRow) return fromRow;
    const y = normStr(row.year);
    const yy = /^\d{4}$/.test(y) ? y : '';
    const tail = String(normCode(row.code) || '')
        .replace(/[^0-9A-Za-z]/g, '')
        .toLowerCase()
        .slice(0, 24);
    if (yy && tail) return ('jg' + yy + tail).toLowerCase().slice(0, 60);
    if (tail) return ('jg' + tail).toLowerCase().slice(0, 60);
    return '';
}
