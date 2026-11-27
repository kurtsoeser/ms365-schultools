/**
 * Öffentliche Intranet-Seite der Schule.
 * Die Schul-IT hinterlegt die Adresse im Intranet-Hub.
 * Lehrkräfte und Schüler/innen öffnen genau diese Seite, nicht den Einrichten-Assistenten.
 */

/**
 * @param {unknown} raw
 * @returns {string}
 */
export function normalizeHttpsUrl(raw) {
    const value = String(raw || '').trim();
    if (!value) return '';
    let url;
    try {
        url = new URL(value);
    } catch {
        return '';
    }
    if (url.protocol !== 'https:') return '';
    if (!url.hostname) return '';
    return url.href;
}

/**
 * Öffentliche Hub-Seite zuerst, sonst die Intranet-Site aus den Stammdaten.
 * @param {object|null|undefined} setup
 * @returns {string}
 */
export function resolveSchoolIntranetOpenUrl(setup) {
    const s = setup && typeof setup === 'object' ? setup : {};
    return normalizeHttpsUrl(s.schoolIntranetSiteUrl) || normalizeHttpsUrl(s.intranetSiteUrl);
}

/**
 * @returns {string}
 */
export function readSchoolIntranetOpenUrl() {
    try {
        if (typeof window === 'undefined' || !window.ms365AppDataV2) return '';
        const setup =
            typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        return resolveSchoolIntranetOpenUrl(setup);
    } catch {
        return '';
    }
}

/**
 * Wer die Hub-Seite öffnet.
 * Schul-IT bleibt im Assistenten. Alle anderen gehen auf die hinterlegte SharePoint-Seite.
 * @param {{ loggedIn?: boolean, isIt?: boolean }|null|undefined} personas
 * @param {object|null|undefined} setup
 * @returns {{ action: 'wait'|'stay'|'open'|'missing', url?: string }}
 */
export function intranetHubAudienceDecision(personas, setup) {
    if (!personas || !personas.loggedIn) return { action: 'wait' };
    if (personas.isIt) return { action: 'stay' };
    const url = resolveSchoolIntranetOpenUrl(setup);
    if (url) return { action: 'open', url };
    return { action: 'missing' };
}
