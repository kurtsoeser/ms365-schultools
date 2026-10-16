/**
 * Materializer: nach Stammdaten-/WebUntis-/Belegungs-Änderungen Landkarte nachziehen.
 * Kein manueller „Abgleich Listen ↔ Landkarte“ im Alltag – läuft mit jedem Schreibpfad.
 * Kanon: core.schulOrganigramm bleibt mit Listen + unterrichtsbelegung synchron.
 */
import {
    syncSchulOrganigrammWithStammdaten,
    normalizeSchulOrganigramm,
    mergeSchulOrganigrammLivePositions
} from '../tools/schul-organigramm/schul-organigramm-logic.js';

let syncing = false;

/** @deprecated Name; nutzen materializeSchulOrganigrammFromAppData / syncSchulOrganigrammFromAppData */
export const MATERIALIZE_ALIAS = 'syncSchulOrganigrammFromAppData';

/**
 * Settings-Objekt für Seed/Sync aus App-Data-v2-Container bauen.
 * @param {object} container
 * @param {object|null|undefined} [belegungOverride]
 */
export function settingsFromAppContainer(container, belegungOverride) {
    const c = container && typeof container === 'object' ? container : {};
    const core = c.core && typeof c.core === 'object' ? c.core : {};
    const years = c.years && typeof c.years === 'object' ? c.years : {};
    const cur = String(years.current || '').trim();
    const bucket =
        (cur && years.byLabel && years.byLabel[cur]) ||
        (years.byLabel && typeof years.byLabel === 'object'
            ? years.byLabel[Object.keys(years.byLabel)[0]]
            : null) ||
        {};
    const belegung =
        belegungOverride != null
            ? belegungOverride
            : bucket.unterrichtsbelegung != null
              ? bucket.unterrichtsbelegung
              : null;
    return {
        schoolName: core.schoolName || '',
        domain: core.domain || '',
        subjects: Array.isArray(core.subjects) ? core.subjects : [],
        arges: Array.isArray(core.arges) ? core.arges : [],
        teachers: Array.isArray(core.teachers) ? core.teachers : [],
        admin: Array.isArray(core.admin) ? core.admin : [],
        adminRoles: Array.isArray(core.adminRoles) ? core.adminRoles : [],
        administration: Array.isArray(core.administration) ? core.administration : [],
        verwaltungAudienceGroups: Array.isArray(core.verwaltungAudienceGroups)
            ? core.verwaltungAudienceGroups
            : [],
        adminAudienceMemberships: Array.isArray(core.adminAudienceMemberships)
            ? core.adminAudienceMemberships
            : [],
        sga: Array.isArray(core.sga) ? core.sga : [],
        sgaMode: core.sgaMode || 'group',
        classes: Array.isArray(bucket.classes) ? bucket.classes : [],
        students: Array.isArray(bucket.students) ? bucket.students : [],
        studentCouncil: Array.isArray(bucket.studentCouncil) ? bucket.studentCouncil : [],
        unterrichtsbelegung: belegung,
        schulOrganigramm: core.schulOrganigramm || { version: 2, nodes: [] }
    };
}

/**
 * Organigramm mit aktuellen Stammdaten + Belegung materialisieren und speichern.
 * @param {{ reason?: string, belegung?: object|null, quiet?: boolean }} [opts]
 * @returns {{ ok: boolean, skipped?: boolean, added?: number, removed?: number, healed?: boolean, error?: string }}
 */
export function syncSchulOrganigrammFromAppData(opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    if (syncing) return { ok: false, skipped: true };
    const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
    if (!api || typeof api.getContainer !== 'function') {
        return { ok: false, error: 'App-Daten nicht verfügbar.' };
    }

    syncing = true;
    if (typeof window !== 'undefined') window.__ms365SchulOrgSyncing = true;
    try {
        const container = api.getContainer();
        if (!container || !container.core) {
            return { ok: false, error: 'Container fehlt.' };
        }
        if (typeof window !== 'undefined' && window.__ms365SchulOrgSkipSyncOnce) {
            window.__ms365SchulOrgSkipSyncOnce = false;
            return { ok: true, skipped: true, reason: o.reason || 'skip-once' };
        }
        const settings = settingsFromAppContainer(container, o.belegung);
        const prev = container.core.schulOrganigramm;
        const result = syncSchulOrganigrammWithStammdaten(prev, settings);
        // Root-Orte/Größe aus dem vorherigen Layout behalten (Speichern darf Landkarte nicht zerlegen)
        const nextLayout = mergeSchulOrganigrammLivePositions(
            normalizeSchulOrganigramm(result.layout),
            prev
        );
        container.core.schulOrganigramm = nextLayout;

        if (typeof api.setContainer === 'function') {
            api.setContainer(container);
        }

        return {
            ok: true,
            added: result.added || 0,
            removed: result.removed || 0,
            healed: !!result.healed,
            labelsUpdated: result.labelsUpdated || 0,
            reason: o.reason || ''
        };
    } catch (e) {
        return {
            ok: false,
            error: e && e.message ? String(e.message) : String(e)
        };
    } finally {
        syncing = false;
        if (typeof window !== 'undefined') window.__ms365SchulOrgSyncing = false;
    }
}

/** Alias für Klarheit im Lokal-First-Modell (Listen → Landkarte). */
export function materializeSchulOrganigrammFromAppData(opts) {
    return syncSchulOrganigrammFromAppData(opts);
}

function installHook() {
    if (typeof window === 'undefined') return;
    window.ms365SyncSchulOrganigrammFromAppData = syncSchulOrganigrammFromAppData;
    window.ms365MaterializeSchulOrganigrammFromAppData = materializeSchulOrganigrammFromAppData;
}

installHook();

export default {
    settingsFromAppContainer,
    syncSchulOrganigrammFromAppData,
    materializeSchulOrganigrammFromAppData
};
