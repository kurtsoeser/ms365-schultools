/**
 * Schulprofil ↔ Lizenz-Admin: Prefill aus /license/me und Rückschreiben (IT-Kontakt).
 */
import { normalizeSchoolProfile } from './school-profile-logic.js';
import {
    buildLicenseProfilePatchFromLocal,
    licenseProfilePatchIsEmpty,
    mergeLicenseIntoLocalFillEmpty,
    normalizeLicenseProfileSnapshot
} from './school-profile-license-sync-logic.js';
import { notifyAppLocalDataChanged } from './app-local-data-notify.js';

const PULL_DONE_KEY = 'ms365-school-profile-license-pull-v1';
const LAST_PUSH_KEY = 'ms365-school-profile-license-push-v1';
const MIN_PUSH_GAP_MS = 8 * 1000;

let pushTimer = null;
/** Verhindert parallele Pulls (z. B. doppelter Mount / Redirect-Wiederholung). */
let pullInFlight = null;
let pushInFlight = null;

function toast(message, kind) {
    try {
        if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(message, {
                kind: kind || 'ok',
                title: 'Schulprofil'
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

function readCoreIdentity() {
    const nameEl = document.getElementById('schoolName');
    const domainEl = document.getElementById('schoolEmailDomain');
    const fromForm = {
        schoolName: nameEl ? String(nameEl.value || '').trim() : '',
        domain: domainEl ? String(domainEl.value || '').trim() : ''
    };
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            return {
                schoolName: fromForm.schoolName || String(s.schoolName || '').trim(),
                domain: fromForm.domain || String(s.domain || '').trim()
            };
        }
    } catch {
        /* ignore */
    }
    try {
        const api = window.ms365AppDataV2;
        if (api && typeof api.getContainer === 'function') {
            const c = api.getContainer();
            const core = c && c.core ? c.core : {};
            return {
                schoolName: fromForm.schoolName || String(core.schoolName || '').trim(),
                domain: fromForm.domain || String(core.domain || '').trim()
            };
        }
    } catch {
        /* ignore */
    }
    return fromForm;
}

function readLocalProfile() {
    try {
        const setup =
            window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function'
                ? window.ms365AppDataV2.getSetup()
                : null;
        return normalizeSchoolProfile(setup && setup.schoolProfile);
    } catch {
        return normalizeSchoolProfile(null);
    }
}

function writeLocalFromMerge(merged) {
    if (!merged || !merged.changed) return false;
    try {
        if (merged.schoolName || merged.domain) {
            if (typeof window.ms365TenantSettingsLoad === 'function') {
                const cur = window.ms365TenantSettingsLoad() || {};
                const next = Object.assign({}, cur, {
                    schoolName: merged.schoolName || cur.schoolName || '',
                    domain: merged.domain || cur.domain || ''
                });
                if (typeof window.ms365TenantSettingsSave === 'function') {
                    window.ms365TenantSettingsSave(next);
                } else if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchCore === 'function') {
                    window.ms365AppDataV2.patchCore({
                        schoolName: next.schoolName,
                        domain: next.domain
                    });
                }
            } else if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchCore === 'function') {
                window.ms365AppDataV2.patchCore({
                    schoolName: merged.schoolName,
                    domain: merged.domain
                });
            }
        }
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({
                schoolProfile: normalizeSchoolProfile(merged.profile)
            });
        }
        notifyAppLocalDataChanged('school-profile-license-pull');
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-tenant-settings-changed', {
                    detail: { reason: 'license-pull', source: 'school-profile-license-sync' }
                })
            );
        } catch {
            /* ignore */
        }
        return true;
    } catch {
        return false;
    }
}

/**
 * Hintergrund: nur Silent-Token – nie Redirect/Popup (sonst Lade-Schleife auf tenant.html).
 * @returns {Promise<string|null>}
 */
async function acquireLicenseTokenSilent() {
    const api = typeof window !== 'undefined' ? window.ms365LicenseApi : null;
    if (!api || typeof api.acquireLicenseToken !== 'function') return null;
    try {
        return await api.acquireLicenseToken({ silentOnly: true });
    } catch {
        return null;
    }
}

/**
 * Leere lokale Felder aus Lizenz-Admin füllen (einmal pro Session / force).
 * @param {{ force?: boolean, toast?: boolean }} [opts]
 */
export async function pullSchoolProfileFromLicense(opts) {
    const force = !!(opts && opts.force);
    try {
        if (!force && sessionStorage.getItem(PULL_DONE_KEY) === '1') {
            return { ok: true, skipped: true, reason: 'already_pulled' };
        }
    } catch {
        /* ignore */
    }

    if (pullInFlight) return pullInFlight;

    const api = typeof window !== 'undefined' ? window.ms365LicenseApi : null;
    if (!api || typeof api.fetchLicenseMe !== 'function') {
        return { ok: false, reason: 'api_missing' };
    }

    pullInFlight = (async function () {
        try {
            const token = await acquireLicenseTokenSilent();
            if (!token) {
                // Kein Silent-Token: nicht redirecten – später nach Login erneut versuchen
                return { ok: false, reason: 'silent_unavailable' };
            }
            const data = await api.fetchLicenseMe(token);
            if (!data || !data.allowed || !data.license) {
                try {
                    sessionStorage.setItem(PULL_DONE_KEY, '1');
                } catch {
                    /* ignore */
                }
                return { ok: false, reason: (data && data.reason) || 'license_denied' };
            }
            const snap = normalizeLicenseProfileSnapshot({
                schoolName: data.license.schoolName,
                primaryDomain: data.license.primaryDomain,
                schoolStammdaten: data.license.schoolStammdaten
            });
            const identity = readCoreIdentity();
            const merged = mergeLicenseIntoLocalFillEmpty(
                { schoolName: identity.schoolName, domain: identity.domain, profile: readLocalProfile() },
                snap
            );
            const wrote = writeLocalFromMerge(merged);
            try {
                sessionStorage.setItem(PULL_DONE_KEY, '1');
            } catch {
                /* ignore */
            }
            // Prefill still – kein Banner/Toast; Formular wird vom Mount neu geladen.
            return { ok: true, changed: wrote, merged, isDesignatedSchoolIt: !!data.isDesignatedSchoolIt };
        } catch (e) {
            return { ok: false, reason: 'error', error: e };
        } finally {
            pullInFlight = null;
        }
    })();

    return pullInFlight;
}

function recentlyPushed(fingerprint) {
    try {
        const raw = localStorage.getItem(LAST_PUSH_KEY) || '';
        const parts = raw.split('|');
        if (parts[0] !== fingerprint) return false;
        const at = Number(parts[1] || 0);
        return Number.isFinite(at) && Date.now() - at < MIN_PUSH_GAP_MS;
    } catch {
        return false;
    }
}

function markPushed(fingerprint) {
    try {
        localStorage.setItem(LAST_PUSH_KEY, fingerprint + '|' + Date.now());
    } catch {
        /* ignore */
    }
}

/**
 * Aktuelles lokales Profil an Lizenz-Admin melden (nur nicht-leere Sync-Felder).
 * @param {{ toast?: boolean, force?: boolean }} [opts]
 */
export async function pushSchoolProfileToLicense(opts) {
    const identity = readCoreIdentity();
    const profile = readLocalProfile();
    const patch = buildLicenseProfilePatchFromLocal({
        schoolName: identity.schoolName,
        profile
    });
    if (licenseProfilePatchIsEmpty(patch)) {
        return { ok: true, skipped: true, reason: 'empty_patch' };
    }

    const fingerprint = JSON.stringify(patch);
    if (!(opts && opts.force) && recentlyPushed(fingerprint)) {
        return { ok: true, skipped: true, reason: 'recent_same' };
    }

    const api = typeof window !== 'undefined' ? window.ms365LicenseApi : null;
    if (!api || typeof api.reportSchoolProfile !== 'function') {
        return { ok: false, reason: 'api_missing' };
    }

    if (pushInFlight) return pushInFlight;

    pushInFlight = (async function () {
        try {
            // Hintergrund-Push: nur Silent – kein Auth-Redirect bei Autosave
            const token = await acquireLicenseTokenSilent();
            if (!token) {
                return { ok: false, reason: 'silent_unavailable' };
            }
            const result = await api.reportSchoolProfile(token, patch);
            if (result && result.ok) {
                markPushed(fingerprint);
                // Erfolgsmeldung kommt vom Speichern-Button („Speichern erfolgreich“), nicht hier.
                return { ok: true, result };
            }
            if (opts && opts.toast) {
                toast(
                    (result && result.message) ||
                        'Schulprofil lokal gespeichert – Meldung an den Betreiber gerade nicht möglich.',
                    'warn'
                );
            }
            return { ok: false, reason: (result && result.reason) || 'rejected', result };
        } catch (e) {
            const status = e && e.status;
            if (opts && opts.toast) {
                const msg =
                    status === 403
                        ? 'Schulprofil lokal gespeichert. Melden an den Betreiber nur für hinterlegte IT-Kontakte.'
                        : status === 401
                          ? 'Schulprofil lokal gespeichert. Für die Betreiber-Meldung bitte anmelden.'
                          : 'Schulprofil lokal gespeichert – Betreiber-Meldung fehlgeschlagen.';
                toast(msg, 'warn');
            }
            return { ok: false, reason: 'error', error: e };
        } finally {
            pushInFlight = null;
        }
    })();

    return pushInFlight;
}

/**
 * Debounced Push nach lokalen Profiländerungen.
 * @param {{ toast?: boolean }} [opts]
 */
export function schedulePushSchoolProfileToLicense(opts) {
    if (pushTimer) clearTimeout(pushTimer);
    pushTimer = setTimeout(function () {
        pushTimer = null;
        pushSchoolProfileToLicense(opts || { toast: false }).catch(function () {
            /* ignore */
        });
    }, 1200);
}

/**
 * Früher Status-Banner – bewusst entfernt (Prefill still, Feedback über Speichern).
 * @param {{ tone?: string, text?: string }} [_opts]
 */
export function paintSchoolProfileLicenseStatus(_opts) {
    const el = typeof document !== 'undefined' ? document.getElementById('tenantSchoolProfileLicenseStatus') : null;
    if (el) el.hidden = true;
}

/**
 * Beim Öffnen des Schulprofils: leere Felder still aus der Lizenz vorfüllen, Formular neu laden.
 * @param {{ reloadForm?: () => void }} [hooks]
 */
export async function mountSchoolProfileLicenseSync(hooks) {
    paintSchoolProfileLicenseStatus();
    const pull = await pullSchoolProfileFromLicense({ toast: false });
    if (pull && pull.changed && hooks && typeof hooks.reloadForm === 'function') {
        hooks.reloadForm();
    }
    return pull;
}

if (typeof window !== 'undefined') {
    window.ms365SchoolProfileLicenseSync = {
        pull: pullSchoolProfileFromLicense,
        push: pushSchoolProfileToLicense,
        schedulePush: schedulePushSchoolProfileToLicense,
        paintStatus: paintSchoolProfileLicenseStatus,
        mount: mountSchoolProfileLicenseSync
    };
}
