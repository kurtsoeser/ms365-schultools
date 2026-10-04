/**
 * Fach-Kürzel: Basis + Varianten (Endziffer, Ü, Plus).
 * Gemeinsam für Kursteams, WebUntis-Import und Stammdaten-Bereinigung.
 */

const SUBJECT_LETTER = /[A-ZÄÖÜ]/;

export function normalizeSubjectToken(s) {
    return String(s || '').trim().toUpperCase();
}

/**
 * @returns {{ base: string, suffix: string }}
 */
export function splitSubjectBaseAndSuffix(token) {
    const t = normalizeSubjectToken(token);
    if (!t) return { base: '', suffix: '' };

    const digitM = t.match(/^(.+?)(\d+)$/);
    if (digitM && digitM[1] && SUBJECT_LETTER.test(digitM[1])) {
        return { base: digitM[1], suffix: digitM[2] };
    }

    if (t.length >= 2 && t.endsWith('Ü')) {
        const base = t.slice(0, -1);
        if (base && SUBJECT_LETTER.test(base)) {
            return { base, suffix: 'Ü' };
        }
    }

    const plus = t.indexOf('+');
    if (plus > 0) {
        const base = t.slice(0, plus);
        const suffix = t.slice(plus);
        if (base && SUBJECT_LETTER.test(base) && suffix.length >= 1) {
            return { base, suffix };
        }
    }

    return { base: t, suffix: '' };
}

export function subjectBaseToken(token) {
    return splitSubjectBaseAndSuffix(token).base;
}

export function subjectHasVariantSuffix(token) {
    return !!splitSubjectBaseAndSuffix(token).suffix;
}

/**
 * @returns {{ base: string, variants: string[], isFamily: boolean }[]}
 */
export function groupSubjectsByBase(subjects) {
    const map = new Map();
    (subjects || []).forEach(function (raw) {
        const token = normalizeSubjectToken(raw);
        if (!token) return;
        const base = subjectBaseToken(token);
        if (!map.has(base)) map.set(base, new Set());
        map.get(base).add(token);
    });
    const groups = Array.from(map.entries()).map(function ([base, set]) {
        const variants = Array.from(set).sort(function (a, b) {
            return a.localeCompare(b, 'de');
        });
        return {
            base,
            variants,
            isFamily: variants.length > 1 || (variants.length === 1 && variants[0] !== base)
        };
    });
    groups.sort(function (a, b) {
        return a.base.localeCompare(b.base, 'de');
    });
    return groups;
}
