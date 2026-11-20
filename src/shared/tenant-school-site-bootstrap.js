/**
 * Tenant → Schul-SharePoint-Zeiger (aus Lizenz) + Bootstrap-Pull der Stammdaten.
 *
 * Betreiber speichert nur die Site-URL (kein Klassen-/Lehrer-Inhalt).
 * Der eingeloggte Schul-User liest mit seinem Token aus dem Schul-Tenant.
 */

import {
    writeItLibraryFormDraft,
    loadItMeta,
    saveItMeta,
    downloadCurrentBackup,
    isReady
} from './stammdaten-sharepoint-sync-api.js';
import { tryAutoLinkItLibrary } from './stammdaten-sharepoint-auto-link.js';
import { SCHOOL_TOOL_SITE_KEYS } from './school-tool-sites.js';
import { DEFAULT_FOLDER } from './stammdaten-sharepoint-sync-logic.js';

const SESSION_BOOT_KEY = 'ms365-school-site-bootstrap-done-v1';

/**
 * @param {unknown} raw
 * @returns {string}
 */
export function normalizeToolSiteUrl(raw) {
    const s = String(raw || '')
        .trim()
        .replace(/\/+$/, '');
    if (!s) return '';
    try {
        const u = new URL(s);
        if (!/^https?:$/i.test(u.protocol)) return '';
        return (u.origin + u.pathname).replace(/\/+$/, '');
    } catch {
        return '';
    }
}

/**
 * ToolSiteUrl aus License-/me-Antwort.
 * @param {object|null|undefined} result
 */
export function toolSiteUrlFromLicenseResult(result) {
    if (!result || typeof result !== 'object') return '';
    const fromSites =
        result.sites && typeof result.sites === 'object' ? result.sites.toolSiteUrl : '';
    const fromLic =
        result.license && typeof result.license === 'object' ? result.license.toolSiteUrl : '';
    return normalizeToolSiteUrl(fromSites || fromLic || result.toolSiteUrl || '');
}

/**
 * @returns {boolean}
 */
export function isLocalSchoolStammdatenEmpty() {
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad() || {};
            const teachers = Array.isArray(s.teachers) ? s.teachers : [];
            const classes = Array.isArray(s.classes) ? s.classes : [];
            if (teachers.length || classes.length) return false;
        }
    } catch {
        /* ignore */
    }
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getCore === 'function') {
            const core = window.ms365AppDataV2.getCore() || {};
            const teachers = Array.isArray(core.teachers) ? core.teachers : [];
            const classes = Array.isArray(core.classes) ? core.classes : [];
            if (teachers.length || classes.length) return false;
        }
    } catch {
        /* ignore */
    }
    return true;
}

/**
 * Lizenz-Zeiger lokal übernehmen (Setup + IT-Meta + Tool-Site-Keys wenn leer).
 * @param {string} toolSiteUrl
 */
export function applyToolSiteUrlLocally(toolSiteUrl) {
    const url = normalizeToolSiteUrl(toolSiteUrl);
    if (!url) return { applied: false, reason: 'empty' };

    writeItLibraryFormDraft({ siteUrl: url });

    const cur = loadItMeta() || {};
    saveItMeta(
        Object.assign({}, cur, {
            siteUrl: url,
            linkedAt: cur.linkedAt || null
        })
    );

    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            const setup =
                typeof window.ms365AppDataV2.getSetup === 'function'
                    ? window.ms365AppDataV2.getSetup() || {}
                    : {};
            /** @type {Record<string, string>} */
            const patch = {};
            if (!String(setup.intranetSiteUrl || '').trim()) patch.intranetSiteUrl = url;
            if (!String(setup.planerDefaultSiteUrl || '').trim()) patch.planerDefaultSiteUrl = url;
            // Lizenz-Zeiger gewinnt als kanonische Schul-Site, wenn gesetzt
            patch.intranetSiteUrl = url;
            patch.planerDefaultSiteUrl = url;
            window.ms365AppDataV2.patchSetup(patch);
        }
    } catch {
        /* ignore */
    }

    try {
        Object.keys(SCHOOL_TOOL_SITE_KEYS).forEach(function (kind) {
            const key = SCHOOL_TOOL_SITE_KEYS[kind];
            if (!key) return;
            const existing = String(localStorage.getItem(key) || '').trim();
            if (!existing) localStorage.setItem(key, url);
        });
    } catch {
        /* ignore */
    }

    return { applied: true, url: url };
}

function sessionBootKey() {
    try {
        const tid =
            (typeof window.ms365AuthGetTenantId === 'function' && window.ms365AuthGetTenantId()) ||
            '';
        return SESSION_BOOT_KEY + ':' + String(tid || 'x');
    } catch {
        return SESSION_BOOT_KEY;
    }
}

function markBootstrapDone() {
    try {
        sessionStorage.setItem(sessionBootKey(), '1');
    } catch {
        /* ignore */
    }
}

function bootstrapDone() {
    try {
        return sessionStorage.getItem(sessionBootKey()) === '1';
    } catch {
        return false;
    }
}

/**
 * Nach Lizenz-OK: Site-Zeiger setzen und bei leerem Browser Stammdaten aus Schul-SP laden.
 * Nutzt nur den User-Token der Schule – Betreiber sieht die Inhalte nicht.
 *
 * @param {object|null|undefined} licenseResult
 * @param {{ force?: boolean }} [opts]
 */
export async function bootstrapSchoolSiteFromLicense(licenseResult, opts) {
    const options = opts || {};
    if (!licenseResult || licenseResult.allowed !== true) {
        return { skipped: true, reason: 'not-allowed' };
    }
    if (!options.force && bootstrapDone()) {
        return { skipped: true, reason: 'session-done' };
    }

    const url = toolSiteUrlFromLicenseResult(licenseResult);
    if (!url) {
        markBootstrapDone();
        return { skipped: true, reason: 'no-tool-site' };
    }

    applyToolSiteUrlLocally(url);

    try {
        await tryAutoLinkItLibrary({ siteUrl: url });
    } catch {
        /* Bibliothek fehlt ggf. noch – Pull scheitert dann klar */
    }

    if (!isLocalSchoolStammdatenEmpty() && !options.force) {
        markBootstrapDone();
        return { skipped: true, reason: 'local-has-data', url: url };
    }

    if (!isReady()) {
        // Keine IT-Bibliothek: trotzdem Lehrerinnen/Klassen von der Tool-Site versuchen
        try {
            const { ensureLiveSchoolStammdatenForBrowser } = await import(
                './ensure-live-school-stammdaten.js'
            );
            const live = await ensureLiveSchoolStammdatenForBrowser({ force: true });
            markBootstrapDone();
            if (live && live.ok) {
                try {
                    window.dispatchEvent(
                        new CustomEvent('ms365-tenant-settings-changed', {
                            detail: { source: 'license-bootstrap-live-spo' }
                        })
                    );
                } catch {
                    /* ignore */
                }
                return { ok: true, url: url, source: 'live-spo', teacherCount: live.teacherCount };
            }
        } catch {
            /* ignore */
        }
        markBootstrapDone();
        return { skipped: true, reason: 'it-library-not-ready', url: url };
    }

    try {
        await downloadCurrentBackup({
            folder: DEFAULT_FOLDER,
            apply: true,
            recordMeta: true
        });
        markBootstrapDone();
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-tenant-settings-changed', {
                    detail: { source: 'license-bootstrap' }
                })
            );
        } catch {
            /* ignore */
        }
        return { ok: true, url: url };
    } catch (e) {
        try {
            const { ensureLiveSchoolStammdatenForBrowser } = await import(
                './ensure-live-school-stammdaten.js'
            );
            const live = await ensureLiveSchoolStammdatenForBrowser({ force: true });
            markBootstrapDone();
            if (live && live.ok) {
                try {
                    window.dispatchEvent(
                        new CustomEvent('ms365-tenant-settings-changed', {
                            detail: { source: 'license-bootstrap-live-spo' }
                        })
                    );
                } catch {
                    /* ignore */
                }
                return { ok: true, url: url, source: 'live-spo', teacherCount: live.teacherCount };
            }
        } catch {
            /* ignore */
        }
        markBootstrapDone();
        return {
            skipped: true,
            reason: 'pull-failed',
            url: url,
            error: e && e.message ? e.message : String(e)
        };
    }
}

export default {
    normalizeToolSiteUrl,
    toolSiteUrlFromLicenseResult,
    isLocalSchoolStammdatenEmpty,
    applyToolSiteUrlLocally,
    bootstrapSchoolSiteFromLicense
};
