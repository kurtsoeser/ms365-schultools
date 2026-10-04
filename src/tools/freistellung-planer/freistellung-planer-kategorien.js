/**
 * Zusätzliche Freistellungs-Kategorien (Admin) + Merge mit Standard-Liste.
 */
import { KATEGORIE_CHOICES as DEFAULT_KATEGORIE_CHOICES } from './freistellung-planer-schema.js';

export const KATEGORIEN_STORAGE_KEY = 'ms365-freistellung-kategorien-extra-v1';

export { DEFAULT_KATEGORIE_CHOICES as KATEGORIE_CHOICES };

/**
 * @param {unknown} raw
 * @returns {string[]}
 */
export function normalizeExtraKategorien(raw) {
    const arr = Array.isArray(raw) ? raw : [];
    const out = [];
    const seen = new Set();
    arr.forEach((x) => {
        const s = String(x || '').trim();
        if (!s || s.length < 2) return;
        const key = s.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(s);
    });
    return out;
}

/**
 * @param {string[]} [extra]
 * @returns {string[]}
 */
export function mergeKategorieChoices(extra) {
    const out = [];
    const seen = new Set();
    DEFAULT_KATEGORIE_CHOICES.forEach((k) => {
        const key = k.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(k);
    });
    normalizeExtraKategorien(extra).forEach((k) => {
        const key = k.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(k);
    });
    return out;
}

export function loadExtraKategorien() {
    try {
        const raw = JSON.parse(localStorage.getItem(KATEGORIEN_STORAGE_KEY) || '[]');
        return normalizeExtraKategorien(raw);
    } catch {
        return [];
    }
}

/**
 * @param {string[]} extra
 */
export function saveExtraKategorien(extra) {
    const next = normalizeExtraKategorien(extra);
    try {
        localStorage.setItem(KATEGORIEN_STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/**
 * @param {string} label
 * @param {string[]} [current]
 */
export function addExtraKategorie(label, current) {
    const list = normalizeExtraKategorien(current || loadExtraKategorien());
    const s = String(label || '').trim();
    if (!s) return list;
    const merged = mergeKategorieChoices(list);
    if (merged.some((k) => k.toLowerCase() === s.toLowerCase())) {
        return list;
    }
    return saveExtraKategorien([...list, s]);
}

/**
 * @param {number} index
 * @param {string[]} [current]
 */
export function removeExtraKategorieAt(index, current) {
    const list = normalizeExtraKategorien(current || loadExtraKategorien());
    const i = Number(index);
    if (i < 0 || i >= list.length) return list;
    list.splice(i, 1);
    return saveExtraKategorien(list);
}

/**
 * @param {string} kategorie
 * @param {string[]} allowed
 */
export function isAllowedKategorie(kategorie, allowed) {
    const kat = String(kategorie || '').trim();
    if (!kat) return false;
    const choices = mergeKategorieChoices(allowed);
    if (choices.some((c) => c === kat)) return true;
    return kat.length >= 2;
}
