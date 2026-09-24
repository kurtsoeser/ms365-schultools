/**
 * Reine sRDP/sRP-Anmeldungs-Logik (kein DOM/Graph) – Vitest.
 */
import {
    VARIANT_LABELS,
    getProfile,
    PROFILE_HAK
} from './srdp-anmeldung-schema.js';

/**
 * @param {string|number|null|undefined} nachname
 * @param {string|number|null|undefined} vorname
 */
export function buildItemTitle(nachname, vorname) {
    const n = String(nachname == null ? '' : nachname).trim();
    const v = String(vorname == null ? '' : vorname).trim();
    if (n && v) return n + ', ' + v;
    return n || v || '';
}

/**
 * @param {string|number|null|undefined} variante
 * @param {string} [profileId]
 */
export function resolveVariant(variante, profileId) {
    const profile = getProfile(profileId);
    const raw = String(variante == null ? '' : variante).trim();
    if (!raw) return null;
    const digit = raw.match(/([123])\s*$/);
    const key = digit ? digit[1] : '';
    const found = profile.variants.find(
        (v) => v.key === key || v.label.toLowerCase() === raw.toLowerCase()
    );
    return found || null;
}

/**
 * @param {string|number|null|undefined} variante
 * @param {string} [profileId]
 */
export function buildPruefplanKurz(variante, profileId) {
    const v = resolveVariant(variante, profileId);
    if (!v) return '';
    return (
        'schriftlich: ' +
        v.schriftlich.join(', ') +
        ' · mündlich: ' +
        v.muendlich.join(', ')
    );
}

/**
 * @param {string|number|null|undefined} wahlfach
 */
export function isSeminarWahlfach(wahlfach) {
    const w = String(wahlfach == null ? '' : wahlfach).trim().toLowerCase();
    if (!w) return false;
    return w === 'seminar' || w.indexOf('seminar') === 0;
}

/**
 * @param {{
 *   profileId?: string,
 *   variante?: string,
 *   wahlfach?: string,
 *   seminar?: string,
 *   lfs?: string,
 *   lehrerLfs?: string,
 *   nachname?: string,
 *   vorname?: string,
 *   bestaetigung?: boolean
 * }} draft
 */
export function validateAnmeldungDraft(draft) {
    const d = draft || {};
    const profile = getProfile(d.profileId);
    /** @type {string[]} */
    const errors = [];
    const variante = resolveVariant(d.variante, profile.id);
    if (!variante) errors.push('Variante fehlt oder ist ungültig.');
    if (!String(d.nachname || '').trim()) errors.push('Nachname fehlt.');
    if (!String(d.vorname || '').trim()) errors.push('Vorname fehlt.');
    if (profile.extraColumnNames.indexOf('LFS') !== -1) {
        if (!String(d.lfs || '').trim()) errors.push((profile.lfsFieldLabel || 'LFS') + ' fehlt.');
        if (!String(d.lehrerLfs || '').trim()) errors.push('LehrerIn LFS fehlt.');
    }
    if (isSeminarWahlfach(d.wahlfach) && !String(d.seminar || '').trim()) {
        errors.push('Seminar-Bezeichnung fehlt (Wahlfach Seminar).');
    }
    if (d.bestaetigung !== true && d.bestaetigung !== 'Ja' && d.bestaetigung !== 'ja') {
        errors.push('Bestätigung der Anmeldung fehlt.');
    }
    return {
        ok: errors.length === 0,
        errors,
        title: buildItemTitle(d.nachname, d.vorname),
        pruefplanKurz: variante ? buildPruefplanKurz(variante.key, profile.id) : '',
        varianteLabel: variante ? variante.label : '',
        profileId: profile.id
    };
}

/**
 * @param {Array<{ code?: string, name?: string, year?: string|number }>} classes
 * @param {string|number} terminJahr
 */
export function filterAbschlussklassen(classes, terminJahr) {
    const y = String(terminJahr || '').trim();
    const list = Array.isArray(classes) ? classes : [];
    if (!/^\d{4}$/.test(y)) return [];
    return list
        .filter((c) => String((c && c.year) || '').trim() === y)
        .map((c) => {
            const code = String((c && c.code) || '').trim();
            const name = String((c && c.name) || '').trim();
            return code || name;
        })
        .filter(Boolean);
}

/**
 * @param {Array<{ code?: string, name?: string }>} subjects
 * @param {string[]} [hints]
 */
export function suggestLfsChoices(subjects, hints) {
    const hintList = Array.isArray(hints) && hints.length ? hints : PROFILE_HAK.lfsHints;
    const fromSubjects = [];
    (Array.isArray(subjects) ? subjects : []).forEach((s) => {
        const code = String((s && s.code) || '').trim().toUpperCase();
        if (!code) return;
        if (/(WS|LFS|EN|FR|IT|SP|RU|LAT)/i.test(code) || hintList.indexOf(code) !== -1) {
            fromSubjects.push(code);
        }
    });
    const seen = new Set();
    const out = [];
    fromSubjects.concat(hintList).forEach((c) => {
        const k = String(c).toUpperCase();
        if (seen.has(k)) return;
        seen.add(k);
        out.push(k);
    });
    return out;
}

/**
 * @param {Array<{ code?: string, name?: string, email?: string }>} teachers
 */
export function teacherChoiceLabels(teachers) {
    const out = [];
    const seen = new Set();
    (Array.isArray(teachers) ? teachers : []).forEach((t) => {
        const name = String((t && t.name) || '').trim();
        const code = String((t && t.code) || '').trim();
        const label = name || code;
        if (!label) return;
        const key = label.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(label);
    });
    out.sort((a, b) => a.localeCompare(b, 'de'));
    return out;
}

/**
 * @param {string} text
 */
export function parseChoiceLines(text) {
    const seen = new Set();
    const out = [];
    String(text || '')
        .split(/\r?\n/)
        .forEach((line) => {
            const s = line.trim();
            if (!s || s.startsWith('#')) return;
            const key = s.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            out.push(s);
        });
    return out;
}

export { VARIANT_LABELS };
export { PROFILE_HAK };
