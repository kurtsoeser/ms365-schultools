/**
 * Optionales Jahrgangs-Filter-Overlay (Freistellungs-Planer only, nicht Dashboard-Persona).
 */
import { normCode } from '../../shared/utils/strings.js';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * @param {unknown} raw
 * @returns {string[]}
 */
export function normalizeAllowedJahrgang(raw) {
    const arr = Array.isArray(raw) ? raw : typeof raw === 'string' ? raw.split(/[,;\s]+/) : [];
    const out = [];
    const seen = new Set();
    arr.forEach((x) => {
        const s = String(x || '').trim();
        if (!s || seen.has(s)) return;
        seen.add(s);
        out.push(s);
    });
    return out;
}

/**
 * @param {unknown} raw
 * @returns {Array<{ jahrgang: string, groupId: string, groupLabel: string }>}
 */
export function normalizeJahrgangGroups(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    const out = [];
    arr.forEach((entry) => {
        const o = entry && typeof entry === 'object' ? entry : {};
        const jahrgang = String(o.jahrgang || o.stufe || '').trim();
        const groupId = String(o.groupId || o.id || '').trim();
        const groupLabel = String(o.groupLabel || o.label || o.group || '').trim();
        if (!jahrgang || !groupId || !GUID_RE.test(groupId)) return;
        out.push({ jahrgang, groupId, groupLabel });
    });
    return out;
}

/**
 * Schulstufe aus Klassenkürzel (z. B. 5A → 5, 10B → 10).
 * @param {string} classCode
 */
export function schulstufeFromClassCode(classCode) {
    const code = normCode(classCode);
    if (!code) return '';
    const m = /^(\d+)/.exec(code);
    return m ? m[1] : '';
}

/**
 * @param {{ allowedJahrgang?: string[], jahrgangGroups?: object[] }} config
 * @returns {string[]}
 */
export function listJahrgangEntraGroupIds(config) {
    const groups = normalizeJahrgangGroups(config && config.jahrgangGroups);
    const ids = groups.map((g) => g.groupId).filter((id) => GUID_RE.test(id));
    return [...new Set(ids)];
}

/**
 * @param {Set<string>|string[]} memberIds
 * @param {{ allowedJahrgang?: string[], jahrgangGroups?: object[] }} config
 * @returns {string[]}
 */
export function jahrgangeFromEntraMembership(memberIds, config) {
    const groups = normalizeJahrgangGroups(config && config.jahrgangGroups);
    if (!groups.length) return [];
    const member = new Set();
    if (memberIds instanceof Set) {
        memberIds.forEach((id) => member.add(String(id).toLowerCase()));
    } else {
        (memberIds || []).forEach((id) => member.add(String(id).toLowerCase()));
    }
    const allow = normalizeAllowedJahrgang(config && config.allowedJahrgang);
    const allowSet = allow.length ? new Set(allow) : null;
    const hit = [];
    const seen = new Set();
    groups.forEach((g) => {
        const gid = String(g.groupId || '').toLowerCase();
        if (!member.has(gid)) return;
        const jg = String(g.jahrgang || '').trim();
        if (!jg || seen.has(jg)) return;
        if (allowSet && !allowSet.has(jg)) return;
        seen.add(jg);
        hit.push(jg);
    });
    return hit.sort();
}

/**
 * @param {string[]} jahrgange
 * @param {object[]} classes
 * @returns {Set<string>}
 */
export function classCodesForJahrgange(jahrgange, classes) {
    const jgSet = new Set((jahrgange || []).map((j) => String(j).trim()).filter(Boolean));
    const codes = new Set();
    if (!jgSet.size) return codes;
    (classes || []).forEach((row) => {
        const code = normCode(row && row.code);
        if (!code) return;
        const stufe = schulstufeFromClassCode(code);
        if (stufe && jgSet.has(stufe)) codes.add(code);
    });
    return codes;
}

/**
 * @param {{ allowedJahrgang?: string[], jahrgangGroups?: object[] }} config
 * @param {Set<string>|string[]|undefined} memberIds
 * @param {object[]} [classes]
 * @returns {{ jahrgange: string[], classCodes: Set<string> }|null}
 */
export function buildJahrgangScope(config, memberIds, classes) {
    const jahrgange = jahrgangeFromEntraMembership(memberIds, config || {});
    if (!jahrgange.length) return null;
    const classCodes = classCodesForJahrgange(jahrgange, classes);
    return { jahrgange, classCodes };
}

/**
 * @param {Set<string>|undefined} jahrgangClassCodes
 * @param {string} itemKlasse
 */
export function itemMatchesJahrgangClassCodes(jahrgangClassCodes, itemKlasse) {
    if (!jahrgangClassCodes || !jahrgangClassCodes.size) return true;
    const code = normCode(itemKlasse);
    return code ? jahrgangClassCodes.has(code) : false;
}
