/**
 * Öffentliche Config laden: window → ?schule= → Demo-Default.
 * Enthält keine Secrets (kein PIN-Pepper, keine Webhook-Secrets).
 */

import { resolveFieldProfile } from './stammdaten-portal-fields.js';

const DEFAULT_CONFIG = Object.freeze({
    schoolName: 'MS365-Schul-Tools Demo',
    schoolSlug: 'demo',
    mode: 'local',
    apiBaseUrl: '',
    presetId: 'kern',
    fieldOverrides: {},
    privacyUrl: ''
});

/**
 * @param {{ search?: string, win?: Window }} [opts]
 */
export async function loadPublicPortalConfig(opts = {}) {
    const win = opts.win || (typeof window !== 'undefined' ? window : null);
    const search = opts.search ?? (win ? win.location.search : '');
    const params = new URLSearchParams(search);

    if (win?.MS365_STAMMDATEN_PORTAL && typeof win.MS365_STAMMDATEN_PORTAL === 'object') {
        return normalizeConfig(win.MS365_STAMMDATEN_PORTAL, params);
    }

    const slug = String(params.get('schule') || '').trim();
    if (slug && win?.location?.href) {
        const candidates = [
            new URL(`../stammdaten-portal-tenants/${encodeURIComponent(slug)}.json`, win.location.href).href,
            new URL(`/stammdaten-portal-tenants/${encodeURIComponent(slug)}.json`, win.location.origin).href
        ];
        for (const href of candidates) {
            try {
                const res = await fetch(href, { cache: 'no-store' });
                if (res.ok) {
                    const json = await res.json();
                    return normalizeConfig(json, params);
                }
            } catch {
                /* try next */
            }
        }
    }

    const modeParam = String(params.get('mode') || '').trim();
    return normalizeConfig({ ...DEFAULT_CONFIG, mode: modeParam || 'local' }, params);
}

function normalizeConfig(raw, params) {
    const modeParam = String(params.get('mode') || '').trim();
    const apiParam = String(params.get('api') || '').trim();
    const cfg = {
        ...DEFAULT_CONFIG,
        ...(raw || {}),
        mode: modeParam || raw?.mode || DEFAULT_CONFIG.mode
    };
    if (apiParam) cfg.apiBaseUrl = apiParam.replace(/\/$/, '');
    // Remote ohne API-URL → zentrales Pilot-Gateway
    if (cfg.mode === 'remote' && !String(cfg.apiBaseUrl || '').trim()) {
        cfg.apiBaseUrl = 'https://func-ms365-ssp-dev.azurewebsites.net/api';
    }
    cfg.profile = resolveFieldProfile(cfg.presetId || 'kern', cfg.fieldOverrides || {});
    return cfg;
}

export { DEFAULT_CONFIG };
