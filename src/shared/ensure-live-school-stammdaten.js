/**
 * Lehrer-Browser: wenn lokale Stammdaten leer sind, Lehrer/Klassen von der Tool-Site lesen
 * und für Portal-Persona + Tools bereitstellen.
 */
import { hydrateStammdatenFromLiveSite } from './school-stammdaten-spo-lists.js';
import {
    loadLiveSchoolStammdatenCache,
    saveLiveSchoolStammdatenCache,
    mergeLiveTeachersIntoLocalTenantIfEmpty,
    loadEffectiveTeachersForPersona,
    isSparseOrDemoOnlyTeacherList
} from './live-school-stammdaten-cache.js';
import { SCHOOL_TOOL_SITE_KEYS } from './school-tool-sites.js';
import {
    normalizeToolSiteUrl,
    toolSiteUrlFromLicenseResult
} from './tenant-school-site-bootstrap.js';

const SESSION_ENSURE_KEY = 'ms365-live-stammdaten-ensure-v1';

/** @type {Promise<{ ok: boolean, teacherCount: number, classCount: number, reason?: string }>|null} */
let inFlight = null;

function markEnsureTried() {
    try {
        sessionStorage.setItem(SESSION_ENSURE_KEY, '1');
    } catch {
        /* ignore */
    }
}

function ensureTriedThisSession() {
    try {
        return sessionStorage.getItem(SESSION_ENSURE_KEY) === '1';
    } catch {
        return false;
    }
}

/** Erneuter SharePoint-Pull erzwingen (z. B. „Klassen neu laden“). */
export function clearLiveStammdatenEnsureFlag() {
    try {
        sessionStorage.removeItem(SESSION_ENSURE_KEY);
    } catch {
        /* ignore */
    }
    inFlight = null;
}

/**
 * @returns {string}
 */
export function resolveToolSiteUrlForLiveStammdaten() {
    const candidates = [];
    // IT-Wahrheit zuerst (Systemadministration / Lizenz) – nicht ein alter LocalStorage-Wert
    // auf Intranet/Root, der SPO „Lehrerinnen“ mit 404 trifft.
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            const setup = window.ms365AppDataV2.getSetup() || {};
            const planerSites =
                setup.planerSiteUrls && typeof setup.planerSiteUrls === 'object'
                    ? setup.planerSiteUrls
                    : {};
            candidates.push(
                planerSites.aktivitaeten,
                setup.planerDefaultSiteUrl,
                toolSiteUrlFromLicenseResult(
                    (window.ms365LicenseGateCore &&
                        typeof window.ms365LicenseGateCore.readCache === 'function' &&
                        (window.ms365LicenseGateCore.readCache() || {}).result) ||
                        null
                )
            );
        }
    } catch {
        /* ignore */
    }
    try {
        const core = window.ms365LicenseGateCore;
        const cache = core && typeof core.readCache === 'function' ? core.readCache() : null;
        candidates.push(toolSiteUrlFromLicenseResult(cache && cache.result));
    } catch {
        /* ignore */
    }
    // Danach Tool-LocalStorage (nur wenn IT/Lizenz noch nichts gesetzt hat)
    try {
        const aktKey = SCHOOL_TOOL_SITE_KEYS && SCHOOL_TOOL_SITE_KEYS.aktivitaeten;
        if (aktKey) candidates.push(localStorage.getItem(aktKey));
    } catch {
        /* ignore */
    }
    try {
        Object.keys(SCHOOL_TOOL_SITE_KEYS || {}).forEach((kind) => {
            if (kind === 'aktivitaeten') return;
            const key = SCHOOL_TOOL_SITE_KEYS[kind];
            if (!key) return;
            candidates.push(localStorage.getItem(key));
        });
    } catch {
        /* ignore */
    }
    try {
        const raw = localStorage.getItem('ms365-it-library-meta-v1');
        if (raw) {
            const meta = JSON.parse(raw);
            if (meta && meta.siteUrl) candidates.push(meta.siteUrl);
        }
    } catch {
        /* ignore */
    }
    // Intranet zuletzt – oft andere Site ohne Lehrerinnen-Liste
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            const setup = window.ms365AppDataV2.getSetup() || {};
            candidates.push(setup.intranetSiteUrl);
        }
    } catch {
        /* ignore */
    }
    for (let i = 0; i < candidates.length; i++) {
        const url = normalizeToolSiteUrl(candidates[i]);
        if (url) return url;
    }
    return '';
}

/**
 * @param {{ force?: boolean }} [opts]
 * @returns {Promise<{ ok: boolean, teacherCount: number, classCount: number, reason?: string, source?: string }>}
 */
export async function ensureLiveSchoolStammdatenForBrowser(opts) {
    const force = !!(opts && opts.force);
    const existing = loadEffectiveTeachersForPersona();
    // 1× „Lehrer01 DEMO“ darf den SharePoint-Vollabzug nicht blockieren
    if (existing.length && !force && !isSparseOrDemoOnlyTeacherList(existing)) {
        return { ok: true, teacherCount: existing.length, classCount: 0, reason: 'local-or-cache' };
    }
    if (!force && ensureTriedThisSession() && !existing.length) {
        const cache = loadLiveSchoolStammdatenCache();
        return {
            ok: !!(cache.teachers && cache.teachers.length),
            teacherCount: (cache.teachers && cache.teachers.length) || 0,
            classCount: (cache.classes && cache.classes.length) || 0,
            reason: 'already-tried'
        };
    }
    if (inFlight && !force) return inFlight;

    inFlight = (async () => {
        const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
        if (!api || typeof api.getGraphToken !== 'function' || typeof api.resolveSiteFromWebUrl !== 'function') {
            markEnsureTried();
            return { ok: false, teacherCount: 0, classCount: 0, reason: 'no-graph' };
        }
        const siteWebUrl = resolveToolSiteUrlForLiveStammdaten();
        if (!siteWebUrl) {
            markEnsureTried();
            return { ok: false, teacherCount: 0, classCount: 0, reason: 'no-site-url' };
        }

        const scopesRead = [
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Sites.Read.All'
        ];
        const scopesRw = [
            'https://graph.microsoft.com/User.Read',
            'https://graph.microsoft.com/Sites.ReadWrite.All'
        ];
        let token = '';
        try {
            if (typeof window.ms365AuthAcquireTokenSilent === 'function') {
                try {
                    token = await window.ms365AuthAcquireTokenSilent(scopesRead);
                } catch {
                    token = await window.ms365AuthAcquireTokenSilent(scopesRw);
                }
            }
        } catch {
            /* popup */
        }
        if (!token) {
            try {
                token = await api.getGraphToken(scopesRead);
            } catch {
                try {
                    token = await api.getGraphToken(scopesRw);
                } catch (e) {
                    markEnsureTried();
                    return {
                        ok: false,
                        teacherCount: 0,
                        classCount: 0,
                        reason: e && e.message ? String(e.message) : 'token-failed'
                    };
                }
            }
        }

        let siteId = '';
        try {
            const site = await api.resolveSiteFromWebUrl(token, siteWebUrl);
            siteId = site && site.id ? String(site.id) : '';
        } catch (e) {
            markEnsureTried();
            return {
                ok: false,
                teacherCount: 0,
                classCount: 0,
                reason: e && e.message ? String(e.message) : 'site-resolve-failed'
            };
        }
        if (!siteId) {
            markEnsureTried();
            return { ok: false, teacherCount: 0, classCount: 0, reason: 'no-site-id' };
        }

        try {
            const result = await hydrateStammdatenFromLiveSite({
                token,
                siteId,
                siteWebUrl,
                graph: api,
                local: { teachers: [], classes: [], students: [] }
            });
            const teachers = (result.stammdaten && result.stammdaten.teachers) || [];
            const classes = (result.stammdaten && result.stammdaten.classes) || [];
            if (teachers.length || classes.length) {
                saveLiveSchoolStammdatenCache({
                    teachers,
                    classes,
                    siteUrl: siteWebUrl,
                    source: (result.live && result.live.source) || 'sharepoint'
                });
                mergeLiveTeachersIntoLocalTenantIfEmpty(teachers, classes);
            }
            markEnsureTried();
            return {
                ok: teachers.length > 0 || classes.length > 0,
                teacherCount: teachers.length,
                classCount: classes.length,
                source: (result.live && result.live.source) || '',
                reason: teachers.length
                    ? 'hydrated'
                    : result.error || 'empty-live'
            };
        } catch (e) {
            markEnsureTried();
            return {
                ok: false,
                teacherCount: 0,
                classCount: 0,
                reason: e && e.message ? String(e.message) : 'hydrate-failed'
            };
        }
    })();

    try {
        return await inFlight;
    } finally {
        inFlight = null;
    }
}
