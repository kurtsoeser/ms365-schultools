/**
 * SharePoint-Berechtigungen für Schulaktivitäten-Paket
 * (Anträge + Regelwerk + Stammdaten Lehrer/Klassen).
 */
import { LIST_TITLES } from './schulaktivitaeten-planer-schema.js';
import {
    defaultStammdatenSpoListTitles,
    resolveStammdatenSpoListTitlesFromWindow
} from '../../shared/school-stammdaten-spo-lists.js';
import { SPO_ROLE } from '../../shared/stammdaten-sharepoint-sync-logic.js';
import { normalizePlannerUsers } from '../freistellung-planer/freistellung-planer-direktion-users.js';
import { overlayStammdatenAudienceOnPermissions } from '../../shared/school-admin-groups.js';
import { notifyAppLocalDataChanged } from '../../shared/app-local-data-notify.js';
import {
    applyEntraGroupListPermissions,
    acquireSpoContext
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';

export const AKT_PERMS_STORAGE_KEY = 'ms365-schulaktivitaeten-perms-v1';

/** @typedef {'read'|'contribute'|'edit'|'fullControl'} PermLevel */

/**
 * Rollen nach Vererbungsbruch.
 * Lehrer: Anträge schreiben, Stammdaten/Regelwerk lesen.
 * Verwaltung: alles steuern (Genehmigung + Regeln + Stammdaten).
 * @type {Record<string, { admin: PermLevel, lehrer: PermLevel|null, schueler: PermLevel|null }>}
 */
export const AKT_LIST_PERM_PROFILES = {
    aktivitaeten: { admin: 'fullControl', lehrer: 'contribute', schueler: null },
    regelwerk: { admin: 'fullControl', lehrer: 'read', schueler: null },
    lehrer: { admin: 'fullControl', lehrer: 'read', schueler: null },
    klassen: { admin: 'fullControl', lehrer: 'read', schueler: null }
};

const DEFAULT_GROUPS = {
    groupAdmin: '',
    groupAdminId: '',
    groupLehrer: '',
    groupLehrerId: '',
    groupSchueler: '',
    groupSchuelerId: ''
};

/**
 * @param {object|null|undefined} raw
 */
export function normalizeAktPermissionsConfig(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    return {
        groupAdmin: String(r.groupAdmin || DEFAULT_GROUPS.groupAdmin).trim(),
        groupAdminId: String(r.groupAdminId || DEFAULT_GROUPS.groupAdminId).trim(),
        groupLehrer: String(r.groupLehrer || DEFAULT_GROUPS.groupLehrer).trim(),
        groupLehrerId: String(r.groupLehrerId || DEFAULT_GROUPS.groupLehrerId).trim(),
        groupSchueler: String(r.groupSchueler || DEFAULT_GROUPS.groupSchueler).trim(),
        groupSchuelerId: String(r.groupSchuelerId || DEFAULT_GROUPS.groupSchuelerId).trim(),
        adminUsers: normalizePlannerUsers(r.adminUsers),
        lehrerUsers: normalizePlannerUsers(r.lehrerUsers),
        schuelerUsers: normalizePlannerUsers(r.schuelerUsers),
        skipPerms: !!r.skipPerms
    };
}

export function loadAktPermissionsConfig() {
    try {
        const raw = JSON.parse(localStorage.getItem(AKT_PERMS_STORAGE_KEY) || '{}');
        return normalizeAktPermissionsConfig(raw);
    } catch {
        return normalizeAktPermissionsConfig({});
    }
}

/** Lehrer-Sammelgruppe aus Stammdaten einbeziehen, falls vorhanden. */
export function loadEffectiveAktPermissionsConfig() {
    return normalizeAktPermissionsConfig(
        overlayStammdatenAudienceOnPermissions(loadAktPermissionsConfig())
    );
}

/**
 * @param {object} patch
 */
export function saveAktPermissionsConfig(patch) {
    const next = normalizeAktPermissionsConfig({
        ...loadAktPermissionsConfig(),
        ...(patch || {})
    });
    try {
        localStorage.setItem(AKT_PERMS_STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    notifyAppLocalDataChanged('schulaktivitaeten-perms');
    return next;
}

export function aktEntraGroupsConfigured(config) {
    const c = normalizeAktPermissionsConfig(config);
    return !!(
        c.groupAdmin ||
        c.groupAdminId ||
        c.groupLehrer ||
        c.groupLehrerId ||
        (c.adminUsers && c.adminUsers.length) ||
        (c.lehrerUsers && c.lehrerUsers.length)
    );
}

/**
 * Titel der Paket-Listen (inkl. Stammdaten-Titel aus Setup).
 * @returns {{ key: string, title: string, profileKey: string }[]}
 */
export function aktPackageListTargets() {
    const stamm = resolveStammdatenSpoListTitlesFromWindow
        ? resolveStammdatenSpoListTitlesFromWindow()
        : defaultStammdatenSpoListTitles();
    return [
        { key: 'aktivitaeten', title: LIST_TITLES.aktivitaeten, profileKey: 'aktivitaeten' },
        { key: 'regelwerk', title: LIST_TITLES.regelwerk, profileKey: 'regelwerk' },
        { key: 'lehrer', title: stamm.lehrer, profileKey: 'lehrer' },
        { key: 'klassen', title: stamm.klassen, profileKey: 'klassen' }
    ];
}

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

async function findListByTitle(token, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

/**
 * Berechtigungen auf alle AKT-Paket-Listen anwenden.
 * @param {string} siteWebUrl
 * @param {object} [configPatch]
 * @param {(msg: string) => void} [logFn]
 */
export async function applySchulaktivitaetenPackagePermissions(siteWebUrl, configPatch, logFn) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const config = normalizeAktPermissionsConfig({
        ...loadAktPermissionsConfig(),
        ...(configPatch || {})
    });
    if (config.skipPerms) {
        write('Berechtigungen übersprungen (skipPerms).');
        return { skipped: true };
    }
    if (!aktEntraGroupsConfigured(config)) {
        write(
            'Berechtigungen: keine Verwaltung/Lehrer-Gruppe oder Einzelperson gewählt – Listen bleiben bei Site-Standardrechten.'
        );
        return { skipped: true, reason: 'no-principals' };
    }
    saveAktPermissionsConfig(config);

    write('Berechtigungen (Verwaltung / Lehrkräfte) …');
    write(
        'Profil: Einreicher → Anträge schreiben, Klassen/Lehrer lesen; Genehmigung → Regeln & Anträge voll.'
    );
    const ctx = await acquireSpoContext(siteWebUrl);
    const site = await G().resolveSiteFromWebUrl(ctx.graphToken, ctx.url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID für Berechtigungen fehlt.');

    const targets = aktPackageListTargets();
    for (let i = 0; i < targets.length; i++) {
        const t = targets[i];
        const list = await findListByTitle(ctx.graphToken, siteId, t.title);
        if (!list) {
            write('  ! „' + t.title + '": Liste nicht gefunden.');
            continue;
        }
        const profile = AKT_LIST_PERM_PROFILES[t.profileKey];
        if (!profile) {
            write('  ! Unbekanntes Profil: ' + t.profileKey);
            continue;
        }
        try {
            await applyEntraGroupListPermissions(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                t.title,
                ctx.graphToken,
                config,
                profile,
                write
            );
        } catch (e) {
            write('  ! „' + t.title + '": ' + (e && e.message ? e.message : e));
        }
        if (typeof G().sleep === 'function') await G().sleep(200);
        else await new Promise((r) => setTimeout(r, 200));
    }
    write('Berechtigungen fertig.');
    return { skipped: false, config };
}

/** @deprecated nur Alias für Tests / Klarheit */
export function aktRoleDefIdForLevel(level) {
    const map = {
        read: SPO_ROLE.read,
        contribute: SPO_ROLE.contribute,
        edit: SPO_ROLE.edit,
        fullControl: SPO_ROLE.fullControl
    };
    return map[level] != null ? map[level] : SPO_ROLE.read;
}
