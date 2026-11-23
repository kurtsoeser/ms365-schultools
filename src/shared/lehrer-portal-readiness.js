/**
 * R13 – „Portal bereit?“ für Lehrkräfte.
 * Pfad A: Schule nicht freigeschaltet → Denied-Hinweis + Landing-Kontakt.
 * Pfad B: freigeschaltet, aber Tool-Site / Listen / IT-Sicherung fehlen.
 *
 * evaluateLehrerPortalReadiness ist rein (ohne window); gather* liest Browser-Stand.
 */

import { toolSiteUrlFromLicenseResult } from './tenant-school-site-bootstrap.js';
import { isReady as isItLibraryReady } from './stammdaten-sharepoint-sync-api.js';
import { loadLiveSchoolStammdatenCache } from './live-school-stammdaten-cache.js';

/** @typedef {'denied'|'setup_incomplete'|'ready'} LehrerPortalReadinessStatus */

/**
 * @typedef {{
 *   status: LehrerPortalReadinessStatus,
 *   missing: string[],
 *   title: string,
 *   message: string
 * }} LehrerPortalReadiness
 */

export const LANDING_KONTAKT_URL = 'https://ms365.schule/#kontakt';

const MISSING_LABELS = {
    license: 'Schul-Lizenz / Freischaltung',
    toolSiteUrl: 'Tool-Site-URL (SharePoint der Schul-Apps)',
    teachers: 'Lehrerliste auf der Tool-Site',
    classes: 'Klassenliste auf der Tool-Site',
    itLibrary: 'IT-Sicherung in der SharePoint-Bibliothek',
    portalRelease: 'IT-Freigabe des Lehrer-Portals'
};

/**
 * @param {string[]} missing
 * @returns {string[]}
 */
export function missingLabels(missing) {
    return (Array.isArray(missing) ? missing : [])
        .map((id) => MISSING_LABELS[id] || String(id || '').trim())
        .filter(Boolean);
}

/**
 * @param {{
 *   allowed?: boolean|null,
 *   toolSiteUrl?: string|null,
 *   teacherCount?: number|null,
 *   classCount?: number|null,
 *   itLibraryReady?: boolean|null
 * }} [input]
 * @returns {LehrerPortalReadiness}
 */
export function evaluateLehrerPortalReadiness(input) {
    const o = input || {};

    if (o.allowed === false) {
        return {
            status: 'denied',
            missing: ['license'],
            title: 'Schule noch nicht freigeschaltet',
            message:
                'Ihre Schule ist für MS365-Schul-Tools noch nicht freigeschaltet. Bitte informieren Sie Ihren IT-Admin oder stellen Sie hier eine Anfrage – wir melden uns.'
        };
    }

    // Ohne abgeschlossene Lizenzprüfung keinen Setup-Banner (vermeidet False Positives).
    if (o.allowed !== true) {
        return {
            status: 'ready',
            missing: [],
            title: '',
            message: ''
        };
    }

    /** @type {string[]} */
    const missing = [];
    const toolSiteUrl = String(o.toolSiteUrl || '').trim();
    if (!toolSiteUrl) missing.push('toolSiteUrl');

    const teacherCount = Math.max(0, Number(o.teacherCount) || 0);
    const classCount = Math.max(0, Number(o.classCount) || 0);
    if (teacherCount <= 0) missing.push('teachers');
    if (classCount <= 0) missing.push('classes');

    // IT-Sicherung nur mitzählen, wenn Stammdaten leer sind (Lehrer-Browser
    // ohne IT-Meta sonst False Positive, obwohl Live-Listen schon greifen).
    if (o.itLibraryReady === false && teacherCount <= 0 && classCount <= 0) {
        missing.push('itLibrary');
    }

    if (!missing.length) {
        return {
            status: 'ready',
            missing: [],
            title: '',
            message: ''
        };
    }

    const labels = missingLabels(missing);
    const list =
        labels.length === 1
            ? labels[0]
            : labels.slice(0, -1).join(', ') + ' und ' + labels[labels.length - 1];

    return {
        status: 'setup_incomplete',
        missing,
        title: 'Einrichtung noch nicht abgeschlossen',
        message:
            'Es fehlen noch Einstellungen an Ihrer Schule (' +
            list +
            '). Bitte weisen Sie Ihren IT-Admin darauf hin, die richtige Tool-Seite zu hinterlegen, die Listen anzulegen bzw. zu befüllen und die IT-Sicherung abzulegen. Danach melden Sie sich neu an.'
    };
}

/**
 * @param {LehrerPortalReadiness | null | undefined} readiness
 * @returns {boolean}
 */
export function isLehrerPortalReady(readiness) {
    return !!(readiness && readiness.status === 'ready');
}

/**
 * Aktuellen Browser-Stand für die Readiness-Prüfung lesen.
 * @returns {{
 *   allowed: boolean|null,
 *   toolSiteUrl: string,
 *   teacherCount: number,
 *   classCount: number,
 *   itLibraryReady: boolean|null
 * }}
 */
export function gatherLehrerPortalReadinessInput() {
    /** @type {{ allowed: boolean|null, toolSiteUrl: string, teacherCount: number, classCount: number, itLibraryReady: boolean|null }} */
    const out = {
        allowed: null,
        toolSiteUrl: '',
        teacherCount: 0,
        classCount: 0,
        itLibraryReady: null
    };
    try {
        const core =
            typeof window !== 'undefined' ? window.ms365LicenseGateCore : null;
        const cache = core && typeof core.readCache === 'function' ? core.readCache() : null;
        if (cache && typeof cache === 'object') {
            if (cache.allowed === true) out.allowed = true;
            else if (cache.allowed === false) out.allowed = false;
            out.toolSiteUrl = toolSiteUrlFromLicenseResult(cache.result || cache) || '';
        }
    } catch {
        /* ignore */
    }
    try {
        if (
            !out.toolSiteUrl &&
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getSetup === 'function'
        ) {
            const setup = window.ms365AppDataV2.getSetup() || {};
            const planer =
                setup.planerSiteUrls && typeof setup.planerSiteUrls === 'object'
                    ? setup.planerSiteUrls
                    : {};
            out.toolSiteUrl = String(
                planer.aktivitaeten || setup.planerDefaultSiteUrl || setup.intranetSiteUrl || ''
            ).trim();
        }
    } catch {
        /* ignore */
    }
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad() || {};
            out.teacherCount = Array.isArray(s.teachers) ? s.teachers.length : 0;
            out.classCount = Array.isArray(s.classes) ? s.classes.length : 0;
        }
    } catch {
        /* ignore */
    }
    try {
        const live = loadLiveSchoolStammdatenCache();
        const liveT = Array.isArray(live.teachers) ? live.teachers.length : 0;
        const liveC = Array.isArray(live.classes) ? live.classes.length : 0;
        if (liveT > out.teacherCount) out.teacherCount = liveT;
        if (liveC > out.classCount) out.classCount = liveC;
        if (!out.toolSiteUrl && live.siteUrl) out.toolSiteUrl = String(live.siteUrl || '').trim();
    } catch {
        /* ignore */
    }
    try {
        out.itLibraryReady = !!isItLibraryReady();
    } catch {
        out.itLibraryReady = null;
    }
    return out;
}

/**
 * @returns {LehrerPortalReadiness}
 */
export function readLehrerPortalReadiness() {
    return evaluateLehrerPortalReadiness(gatherLehrerPortalReadinessInput());
}
