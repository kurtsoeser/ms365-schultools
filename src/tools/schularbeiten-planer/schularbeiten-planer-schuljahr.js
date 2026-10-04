/**
 * Schuljahr-Hilfen (österreichisch: Start 1. September, Format 2026/27).
 */
import { toIsoDateOnly, parseIsoDateParts } from './schularbeiten-planer-logic.js';

/**
 * @param {unknown} input
 * @returns {string} normalisiertes Schuljahr oder ''
 */
export function normalizeSchuljahr(input) {
    const raw = String(input ?? '')
        .trim()
        .replace(/\s+/g, '');
    if (!raw) return '';
    const m = /^(\d{4})[\/.\-](\d{2}|\d{4})$/.exec(raw);
    if (!m) return '';
    let endShort = m[2];
    if (endShort.length === 4) endShort = endShort.slice(2);
    endShort = endShort.padStart(2, '0');
    const startYear = Number(m[1]);
    if (!Number.isFinite(startYear) || startYear < 1990 || startYear > 2100) return '';
    return `${startYear}/${endShort}`;
}

/**
 * @param {Date} [date]
 */
export function currentSchoolYearFromDate(date) {
    const d = date instanceof Date && !Number.isNaN(date.getTime()) ? date : new Date();
    const y = d.getFullYear();
    const month = d.getMonth() + 1;
    const startYear = month >= 9 ? y : y - 1;
    const endShort = String((startYear + 1) % 100).padStart(2, '0');
    return `${startYear}/${endShort}`;
}

/**
 * @param {string} iso
 * @returns {string}
 */
export function inferSchuljahrFromIsoDate(iso) {
    const parts = parseIsoDateParts(toIsoDateOnly(iso) || '');
    if (!parts) return '';
    const { y, m } = parts;
    const startYear = m >= 9 ? y : y - 1;
    const endShort = String((startYear + 1) % 100).padStart(2, '0');
    return `${startYear}/${endShort}`;
}

/**
 * @param {{ schuljahr?: string, datum?: string }} item
 */
export function effectiveSchuljahrForItem(item) {
    const explicit = normalizeSchuljahr(item && item.schuljahr);
    if (explicit) return explicit;
    return inferSchuljahrFromIsoDate(item && item.datum) || '';
}

/**
 * @param {object} item
 * @param {string} selectedSchuljahr leer = alle Jahre
 */
export function matchesSchuljahrFilter(item, selectedSchuljahr) {
    const sel = normalizeSchuljahr(selectedSchuljahr);
    if (!sel) return true;
    return effectiveSchuljahrForItem(item) === sel;
}

/**
 * @param {object[]} items
 * @param {object[]} [windows]
 * @param {object[]} [rules]
 * @param {string} [fallback]
 */
export function collectSchoolYears(items, windows, rules, fallback) {
    const set = new Set();
    const add = (v) => {
        const n = normalizeSchuljahr(v);
        if (n) set.add(n);
    };
    add(fallback);
    add(currentSchoolYearFromDate());
    (items || []).forEach((it) => add(effectiveSchuljahrForItem(it)));
    (windows || []).forEach((w) => add(w.schuljahr));
    (rules || []).forEach((r) => add(r.schuljahr));
    return Array.from(set).sort((a, b) => {
        const ay = Number(String(a).slice(0, 4));
        const by = Number(String(b).slice(0, 4));
        return by - ay;
    });
}

/**
 * @param {object[]} rules
 * @param {string} schuljahr
 */
export function pickActiveRegelwerk(rules, schuljahr) {
    const sj = normalizeSchuljahr(schuljahr);
    const list = Array.isArray(rules) ? rules : [];
    const scoped = list.filter((r) => {
        const rs = normalizeSchuljahr(r.schuljahr);
        if (!rs) return true;
        return !sj || rs === sj;
    });
    return scoped.find((r) => r.aktiv) || scoped[0] || list.find((r) => r.aktiv) || list[0] || null;
}

/**
 * @param {object[]} windows
 * @param {string} schuljahr
 */
export function filterTerminfensterForSchuljahr(windows, schuljahr) {
    const sj = normalizeSchuljahr(schuljahr);
    const list = Array.isArray(windows) ? windows : [];
    if (!sj) return list;
    return list.filter((w) => {
        const ws = normalizeSchuljahr(w.schuljahr);
        if (!ws) return true;
        return ws === sj;
    });
}

/**
 * @param {object[]} fachMeta
 * @param {string} schuljahr
 */
export function filterFachMetaForSchuljahr(fachMeta, schuljahr) {
    const sj = normalizeSchuljahr(schuljahr);
    const list = Array.isArray(fachMeta) ? fachMeta : [];
    if (!sj) return list;
    const scoped = list.filter((m) => {
        const ms = normalizeSchuljahr(m.schuljahr);
        return !ms || ms === sj;
    });
    return scoped.length ? scoped : list;
}
