/**
 * SharePoint-Berechtigungen für Stammdaten-Listen (Entra-Gruppen, wie Schularbeiten-Planer).
 */
import { findListByDisplayName } from '../schularbeiten-planer/schularbeiten-planer-graph.js';
import {
    normalizePermissionsConfig,
    applyEntraGroupListPermissions,
    acquireSpoContext
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';

export const PERMS_STORAGE_KEY = 'ms365-stammdaten-listen-perms-v1';

/** @typedef {'read'|'contribute'|'fullControl'} PermLevel */

/**
 * @type {Record<string, { admin: PermLevel, lehrer: PermLevel|null, schueler: PermLevel|null }>}
 */
export const STAMMDATEN_LIST_PERM_PROFILES = {
    /** Gesamte Schülerliste – für Schüler-Gruppe kein Zugriff (Datenschutz). */
    schueler: { admin: 'fullControl', lehrer: 'contribute', schueler: null },
    faecher: { admin: 'fullControl', lehrer: 'read', schueler: 'read' },
    fachgruppen: { admin: 'fullControl', lehrer: 'read', schueler: 'read' },
    arges: { admin: 'fullControl', lehrer: 'contribute', schueler: null },
    /** Klassen inkl. Personenfeld Schülerinnen: Schüler-Gruppe nur Lesen. */
    klassen: { admin: 'fullControl', lehrer: 'contribute', schueler: 'read' }
};

export const LIST_TYPE_KEYS = ['schueler', 'faecher', 'fachgruppen', 'arges', 'klassen'];

export function loadPermissionsConfig() {
    try {
        const raw = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}');
        return normalizePermissionsConfig(raw);
    } catch {
        return normalizePermissionsConfig({});
    }
}

/**
 * @param {object} patch
 */
export function savePermissionsConfig(patch) {
    const next = normalizePermissionsConfig({ ...loadPermissionsConfig(), ...(patch || {}) });
    try {
        localStorage.setItem(PERMS_STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

/** Übernimmt Gruppen aus Schularbeiten-Setup, wenn hier noch leer. */
export function prefillPermissionsFromSchularbeitenIfEmpty() {
    const cur = loadPermissionsConfig();
    if (cur.groupAdminId || cur.groupAdmin || cur.groupLehrerId || cur.groupLehrer) return cur;
    try {
        const sa = JSON.parse(localStorage.getItem('ms365-schularbeiten-perms-v1') || '{}');
        const merged = normalizePermissionsConfig({
            groupAdmin: sa.groupAdmin,
            groupAdminId: sa.groupAdminId,
            groupLehrer: sa.groupLehrer,
            groupLehrerId: sa.groupLehrerId,
            groupSchueler: sa.groupSchueler,
            groupSchuelerId: sa.groupSchuelerId
        });
        savePermissionsConfig(merged);
        return merged;
    } catch {
        return cur;
    }
}

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

/**
 * @param {string} siteWebUrl
 * @param {object} [configPatch]
 * @param {(msg: string) => void} [logFn]
 * @param {{ schueler?: boolean, faecher?: boolean, fachgruppen?: boolean, arges?: boolean, klassen?: boolean, schuelerTitle?: string, faecherTitle?: string, fachgruppenTitle?: string, argesTitle?: string, klassenTitle?: string, skipPerms?: boolean }} [listOpts]
 */
export async function applyStammdatenPackagePermissions(siteWebUrl, configPatch, logFn, listOpts) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const config = normalizePermissionsConfig({ ...loadPermissionsConfig(), ...(configPatch || {}) });
    const o = listOpts && typeof listOpts === 'object' ? listOpts : {};
    if (config.skipPerms || o.skipPerms) {
        write('Berechtigungen übersprungen.');
        return { skipped: true };
    }
    savePermissionsConfig(config);

    const titles = {
        schueler: String(o.schuelerTitle || 'Schülerinnen').trim() || 'Schülerinnen',
        faecher: String(o.faecherTitle || 'Fächer').trim() || 'Fächer',
        fachgruppen: String(o.fachgruppenTitle || 'Fachgruppen').trim() || 'Fachgruppen',
        arges: String(o.argesTitle || 'ARGEs').trim() || 'ARGEs',
        klassen: String(o.klassenTitle || 'Klassen').trim() || 'Klassen'
    };

    const active = [];
    if (o.schueler) active.push('schueler');
    if (o.faecher) active.push('faecher');
    if (o.fachgruppen) active.push('fachgruppen');
    if (o.arges) active.push('arges');
    if (o.klassen) active.push('klassen');
    if (!active.length) {
        LIST_TYPE_KEYS.forEach((k) => active.push(k));
    }

    write('Berechtigungen (Entra-Gruppen) für Stammdaten-Listen …');
    const ctx = await acquireSpoContext(siteWebUrl);
    const site = await G().resolveSiteFromWebUrl(ctx.graphToken, ctx.url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID für Berechtigungen fehlt.');

    for (let i = 0; i < active.length; i++) {
        const typeKey = active[i];
        const listTitle = titles[typeKey];
        const profile = STAMMDATEN_LIST_PERM_PROFILES[typeKey];
        if (!listTitle || !profile) continue;
        const list = await findListByDisplayName(ctx.graphToken, siteId, listTitle);
        if (!list || !list.id) {
            write('  ! „' + listTitle + '": Liste nicht gefunden – übersprungen.');
            continue;
        }
        const displayTitle = list.displayName ? String(list.displayName) : listTitle;
        try {
            await applyEntraGroupListPermissions(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                displayTitle,
                ctx.graphToken,
                config,
                profile,
                write
            );
        } catch (e) {
            write('  ! „' + displayTitle + '": ' + (e && e.message ? e.message : e));
        }
        await G().sleep(200);
    }
    write('Berechtigungen fertig.');
    return { skipped: false, config };
}
