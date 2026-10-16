/**
 * Ein Schreibpfad für Schul-Stammdaten (Lokal-First).
 *
 * - Schreibt nur in ms365-schooltool-data-v2 (über setCoreFromTenantSettings / ms365TenantSettingsSave)
 * - Materialisiert die Landkarte (schulOrganigramm) im selben Durchlauf
 * - Kein v1-Schreiben, kein SharePoint-Push (Auto-Sync ist lokal-first aus)
 *
 * Tools sollen bevorzugt writeSchoolData / materializeSchulOrganigramm nutzen,
 * nicht direkt parallele localStorage-Keys für dieselbe Domäne.
 */
import { syncSchulOrganigrammFromAppData } from './schul-organigramm-stammdaten-sync.js';

/**
 * Landkarte aus aktuellen Listen nachziehen (kein separater Abgleich-Schritt).
 * @param {{ reason?: string, belegung?: object|null, quiet?: boolean }} [opts]
 */
export function materializeSchulOrganigramm(opts) {
    return syncSchulOrganigrammFromAppData(opts || { reason: 'writeSchoolData' });
}

/**
 * Stammdaten speichern (Listen) und Landkarte atomar materialisieren.
 * @param {object} settings Flat-Settings wie ms365TenantSettingsSave
 * @param {{ reason?: string, skipMaterialize?: boolean }} [opts]
 * @returns {{ settings: object, materialize: object|null }}
 */
export function writeSchoolData(settings, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const reason = String(o.reason || 'writeSchoolData').trim() || 'writeSchoolData';

    if (typeof window === 'undefined' || typeof window.ms365TenantSettingsSave !== 'function') {
        throw new Error('ms365TenantSettingsSave nicht verfügbar.');
    }

    // setCoreFromTenantSettings materialisiert bereits; skipMaterialize nur wenn Save
    // den Organigramm-Hook bewusst umgeht (selten).
    const skipHookOnce = !!o.skipMaterialize;
    if (skipHookOnce) {
        try {
            window.__ms365SchulOrgSkipSyncOnce = true;
        } catch {
            /* ignore */
        }
    }

    const saved = window.ms365TenantSettingsSave(settings);

    let materialize = null;
    if (!skipHookOnce) {
        // Hook in setCoreFromTenantSettings läuft schon; hier nur Status nachziehen,
        // falls der Hook (Modul noch nicht geladen) nichts tat.
        try {
            const api = window.ms365AppDataV2;
            const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
            const nodes =
                c && c.core && c.core.schulOrganigramm && Array.isArray(c.core.schulOrganigramm.nodes)
                    ? c.core.schulOrganigramm.nodes.length
                    : 0;
            materialize = { ok: true, via: 'save-hook', reason: reason, nodeCount: nodes };
        } catch {
            materialize = materializeSchulOrganigramm({ reason: reason });
        }
    } else {
        materialize = { ok: true, skipped: true, reason: reason };
    }

    try {
        window.dispatchEvent(
            new CustomEvent('ms365-school-data-written', {
                detail: { reason: reason, materialize: materialize }
            })
        );
    } catch {
        /* ignore */
    }

    return { settings: saved, materialize: materialize };
}

/**
 * Nur Landkarte speichern (Knoten/Layout), Listen unverändert lassen.
 * Danach optional Listen→Knoten heilen (fehlende Refs nachziehen).
 * @param {object} schulOrganigramm
 * @param {{ reason?: string, materializeLists?: boolean }} [opts]
 */
export function writeSchulOrganigramm(schulOrganigramm, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
    if (!api || typeof api.getContainer !== 'function' || typeof api.setContainer !== 'function') {
        throw new Error('App-Daten (v2) nicht verfügbar.');
    }
    const c = api.getContainer();
    if (!c || !c.core) throw new Error('Container fehlt.');
    c.core.schulOrganigramm = schulOrganigramm;
    // Speichern ohne sofortigen Listen-Materializer (Layout nicht zerlegen)
    try {
        window.__ms365SchulOrgSkipSyncOnce = true;
    } catch {
        /* ignore */
    }
    api.setContainer(c);

    let materialize = null;
    if (o.materializeLists !== false) {
        // Nach Layout-Save: fehlende Knoten aus Listen nachziehen, Orte behalten
        materialize = materializeSchulOrganigramm({
            reason: String(o.reason || 'writeSchulOrganigramm')
        });
    }
    return { ok: true, materialize: materialize };
}

if (typeof window !== 'undefined') {
    window.ms365WriteSchoolData = writeSchoolData;
    window.ms365MaterializeSchulOrganigramm = materializeSchulOrganigramm;
    window.ms365WriteSchulOrganigramm = writeSchulOrganigramm;
    // Alias: bisheriger Sync-Name bleibt, meint jetzt „Materializer“
    if (typeof window.ms365SyncSchulOrganigrammFromAppData !== 'function') {
        window.ms365SyncSchulOrganigrammFromAppData = materializeSchulOrganigramm;
    }
}

export default {
    writeSchoolData,
    writeSchulOrganigramm,
    materializeSchulOrganigramm
};
