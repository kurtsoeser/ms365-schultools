/**
 * Einstellungen Teams-Antrag (Freigeber-Gruppe + Extra-Mails, Nomenklatur).
 */

const KEY = 'ms365-teams-antrag-settings-v1';

/**
 * @returns {{
 *   freigeberMails: string[],
 *   freigeberGroupId: string,
 *   freigeberGroupName: string,
 *   nomenklaturHinweis: string,
 *   schuelerDuertenBeantragen: boolean
 * }}
 */
export function defaultSettings() {
    return {
        freigeberMails: [],
        freigeberGroupId: '',
        freigeberGroupName: '',
        nomenklaturHinweis:
            'Präfixe: ag-, fg-, prj-, kl- + Schuljahr + Kurzname. Keine Umlaute im Alias.',
        schuelerDuertenBeantragen: true
    };
}

/**
 * @param {unknown} raw
 */
export function normalizeSettings(raw) {
    const d = defaultSettings();
    const r = raw && typeof raw === 'object' ? raw : {};
    const mails = Array.isArray(r.freigeberMails)
        ? r.freigeberMails
        : String(r.freigeberMails || '')
              .split(/[\n;,]+/)
              .map((s) => s.trim());
    const groupId = String(r.freigeberGroupId || '').trim();
    return {
        freigeberMails: mails
            .map((s) => String(s || '').trim().toLowerCase())
            .filter((s) => s.includes('@')),
        freigeberGroupId: /^[0-9a-f-]{36}$/i.test(groupId) ? groupId : '',
        freigeberGroupName: String(r.freigeberGroupName || '').trim().slice(0, 200),
        nomenklaturHinweis: String(
            r.nomenklaturHinweis != null ? r.nomenklaturHinweis : d.nomenklaturHinweis
        ),
        schuelerDuertenBeantragen: r.schuelerDuertenBeantragen === false ? false : true
    };
}

export function loadTeamsAntragSettings() {
    try {
        if (typeof localStorage === 'undefined') return defaultSettings();
        const raw = localStorage.getItem(KEY);
        if (!raw) return defaultSettings();
        return normalizeSettings(JSON.parse(raw));
    } catch {
        return defaultSettings();
    }
}

/**
 * @param {Partial<ReturnType<typeof defaultSettings>>} partial
 */
export function saveTeamsAntragSettings(partial) {
    const next = normalizeSettings({ ...loadTeamsAntragSettings(), ...(partial || {}) });
    try {
        if (typeof localStorage !== 'undefined') {
            localStorage.setItem(KEY, JSON.stringify(next));
        }
    } catch {
        /* ignore */
    }
    return next;
}
