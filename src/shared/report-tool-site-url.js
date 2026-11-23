/**
 * Schul-IT → Betreiber: Tool-Site-URL (SharePoint, wo die Listen liegen) an License-API melden.
 * Es wird nur die Adresse gemeldet – keine Listen-/Schüler-/Lehrer-Inhalte.
 */
import { normalizeToolSiteUrl } from './tenant-school-site-bootstrap.js';

const LAST_KEY = 'ms365-reported-tool-site-url-v1';
const LAST_AT_KEY = 'ms365-reported-tool-site-url-at-v1';
const MIN_GAP_MS = 15 * 1000;

/** Kurz für Toast / Status – Datenversprechen. */
export const TOOL_SITE_REPORT_TOAST =
    'Primär-Site-Adresse für den Betrieb hinterlegt. Es wird nur die Webadresse gemeldet – keine Schüler-, Lehrer- oder Listendaten. Die Inhalte bleiben in Ihrem Schul-SharePoint.';

export const TOOL_SITE_REPORT_TOAST_UNCHANGED =
    'Primär-Site-Adresse ist dem Betreiber bereits bekannt. Es liegt weiterhin nur die Webadresse vor – kein Zugriff auf Ihre Schuldaten.';

function showReportToast(message, kind) {
    try {
        if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(message, {
                kind: kind || 'ok',
                title: 'Datenversprechen'
            });
            return;
        }
        if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(message);
        }
    } catch {
        /* ignore */
    }
}

function readLastReported() {
    try {
        return normalizeToolSiteUrl(localStorage.getItem(LAST_KEY) || '');
    } catch {
        return '';
    }
}

function writeLastReported(url) {
    try {
        localStorage.setItem(LAST_KEY, url);
        localStorage.setItem(LAST_AT_KEY, String(Date.now()));
    } catch {
        /* ignore */
    }
}

function recentlyReported(url) {
    const last = readLastReported();
    if (!last || last !== url) return false;
    try {
        const at = Number(localStorage.getItem(LAST_AT_KEY) || 0);
        return Number.isFinite(at) && Date.now() - at < MIN_GAP_MS;
    } catch {
        return false;
    }
}

/**
 * @param {string} toolSiteUrl
 * @param {{ force?: boolean, toast?: boolean, source?: string }} [opts]
 * @returns {Promise<{ ok: boolean, skipped?: boolean, reason?: string, result?: object }>}
 */
export async function reportSchoolToolSiteUrl(toolSiteUrl, opts) {
    const url = normalizeToolSiteUrl(toolSiteUrl);
    if (!url) return { ok: false, reason: 'empty' };

    const force = !!(opts && opts.force);
    if (!force && recentlyReported(url)) {
        return { ok: true, skipped: true, reason: 'recent_same' };
    }
    if (!force && readLastReported() === url) {
        return { ok: true, skipped: true, reason: 'unchanged_local' };
    }

    const api = typeof window !== 'undefined' ? window.ms365LicenseApi : null;
    if (!api || typeof api.acquireLicenseToken !== 'function' || typeof api.reportToolSiteUrl !== 'function') {
        return { ok: false, reason: 'api_missing' };
    }

    try {
        const token = await api.acquireLicenseToken();
        const result = await api.reportToolSiteUrl(token, url);
        if (result && result.ok) {
            writeLastReported(url);
            if (opts && opts.toast) {
                showReportToast(
                    result.unchanged ? TOOL_SITE_REPORT_TOAST_UNCHANGED : TOOL_SITE_REPORT_TOAST,
                    'ok'
                );
            }
            return { ok: true, result };
        }
        if (opts && opts.toast) {
            showReportToast(
                (result && result.message) ||
                    'Primär-Site-Adresse konnte nicht hinterlegt werden. Die Schuldaten bleiben lokal – bitte IT-Kontakt prüfen oder erneut versuchen.',
                'warn'
            );
        }
        return { ok: false, reason: (result && result.reason) || 'rejected', result };
    } catch (e) {
        const status = e && e.status;
        const reason =
            status === 403 ? 'forbidden' : status === 401 ? 'unauthorized' : 'error';
        if (opts && opts.toast) {
            const msg =
                reason === 'forbidden'
                    ? 'Nur hinterlegte IT-Kontakte können die Primär-Site-Adresse melden. Schuldaten wurden nicht übertragen.'
                    : reason === 'unauthorized'
                      ? 'Bitte mit Microsoft 365 anmelden, um die Primär-Site-Adresse zu hinterlegen.'
                      : 'Primär-Site-Adresse konnte nicht gemeldet werden (API ggf. noch nicht aktualisiert). Schuldaten bleiben in Ihrem Tenant.';
            showReportToast(msg, 'warn');
        }
        return {
            ok: false,
            reason,
            result: { message: e && e.message ? String(e.message) : String(e), status }
        };
    }
}

/**
 * Primär-Site aus Stammdaten melden (wenn gesetzt).
 * @param {{ force?: boolean, toast?: boolean }} [opts]
 */
export async function reportPlanerDefaultSiteIfSet(opts) {
    let url = '';
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
            const setup = window.ms365AppDataV2.getSetup() || {};
            url = normalizeToolSiteUrl(setup.planerDefaultSiteUrl || setup.intranetSiteUrl || '');
        }
    } catch {
        /* ignore */
    }
    if (!url) return { ok: false, reason: 'no_local_site' };
    return reportSchoolToolSiteUrl(url, opts);
}

if (typeof window !== 'undefined') {
    window.ms365ReportToolSiteUrl = reportSchoolToolSiteUrl;
    window.ms365ReportPlanerDefaultSite = reportPlanerDefaultSiteIfSet;
}
