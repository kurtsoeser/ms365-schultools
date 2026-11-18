/**
 * Diplomarbeiten / VWA – Betriebsmodus und Typ.
 */

const KEY = 'ms365-diplomarbeiten-settings-v1';

/** @typedef {'diplom'|'vwa'} ArbeitTyp */
/** @typedef {'channels'|'legacy'} Betriebsmodus */

/**
 * @returns {{ arbeitTyp: ArbeitTyp, modus: Betriebsmodus, setupDone: boolean }}
 */
export function defaultDiplomSettings() {
    return {
        arbeitTyp: 'diplom',
        modus: 'channels',
        setupDone: false
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeDiplomSettings(raw) {
    const d = defaultDiplomSettings();
    const r = raw && typeof raw === 'object' ? raw : {};
    const typ = String(r.arbeitTyp || d.arbeitTyp).toLowerCase();
    const modus = String(r.modus || d.modus).toLowerCase();
    return {
        arbeitTyp: typ === 'vwa' ? 'vwa' : 'diplom',
        modus: modus === 'legacy' ? 'legacy' : 'channels',
        setupDone: !!r.setupDone
    };
}

export function loadDiplomSettings() {
    try {
        if (typeof localStorage === 'undefined') return defaultDiplomSettings();
        const raw = localStorage.getItem(KEY);
        if (!raw) return defaultDiplomSettings();
        return normalizeDiplomSettings(JSON.parse(raw));
    } catch {
        return defaultDiplomSettings();
    }
}

/**
 * @param {Partial<ReturnType<typeof defaultDiplomSettings>>} partial
 */
export function saveDiplomSettings(partial) {
    const next = normalizeDiplomSettings({ ...loadDiplomSettings(), ...(partial || {}) });
    try {
        if (typeof localStorage !== 'undefined') {
            localStorage.setItem(KEY, JSON.stringify(next));
        }
    } catch {
        /* ignore */
    }
    return next;
}
