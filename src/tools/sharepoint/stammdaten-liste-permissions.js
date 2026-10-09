/**
 * SharePoint-Berechtigungen für Stammdaten-Listen (Entra-Gruppen, wie Schularbeiten-Planer).
 */
import { findListByDisplayName } from '../schularbeiten-planer/schularbeiten-planer-graph.js';
import { normalizePermissionsConfig, acquireSpoContext } from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import {
    normalizeListPermProfiles,
    normalizeGrantRows,
    migrateLegacyPermConfig,
    mergeAudienceSlotsIntoGrantRows,
    grantsForListKey,
    seedDefaultGrantRows
} from './stammdaten-liste-perm-matrix.js';
import {
    SPO_ROLE,
    entraGroupLogonName,
    isBroadSiteAudience
} from '../../shared/stammdaten-sharepoint-sync-logic.js';
import { overlayStammdatenAudienceOnPermissions } from '../../shared/school-admin-groups.js';

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
    klassen: { admin: 'fullControl', lehrer: 'contribute', schueler: 'read' },
    /** Lehrkräfte-Stammliste – für Schüler-Gruppe kein Zugriff. */
    lehrer: { admin: 'fullControl', lehrer: 'read', schueler: null }
};

export const LIST_TYPE_KEYS = ['schueler', 'faecher', 'fachgruppen', 'arges', 'klassen', 'lehrer'];

function resolveGrantRowsFromRaw(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    let grantRows = normalizeGrantRows(r.grantRows);
    if (!grantRows.length) {
        grantRows = migrateLegacyPermConfig({ ...r, listProfiles: r.listProfiles });
    }
    return grantRows;
}

export function loadPermissionsConfig() {
    try {
        const raw = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}');
        const base = normalizePermissionsConfig(raw);
        const listProfiles = normalizeListPermProfiles(raw.listProfiles);
        const grantRows = resolveGrantRowsFromRaw(raw);
        return { ...base, listProfiles, grantRows };
    } catch {
        const base = normalizePermissionsConfig({});
        return {
            ...base,
            listProfiles: normalizeListPermProfiles(null),
            grantRows: migrateLegacyPermConfig({})
        };
    }
}

export function loadEffectivePermissionsConfig() {
    const cur = loadPermissionsConfig();
    const base = overlayStammdatenAudienceOnPermissions(cur);
    const grantRows = mergeAudienceSlotsIntoGrantRows(cur.grantRows, base);
    return { ...base, listProfiles: cur.listProfiles, grantRows };
}

/**
 * @param {object} patch
 */
export function savePermissionsConfig(patch) {
    const cur = loadPermissionsConfig();
    const next = normalizePermissionsConfig({ ...cur, ...(patch || {}) });
    const listProfiles =
        patch && patch.listProfiles != null
            ? normalizeListPermProfiles(patch.listProfiles)
            : normalizeListPermProfiles(cur.listProfiles);
    const grantRows =
        patch && patch.grantRows != null
            ? normalizeGrantRows(patch.grantRows)
            : normalizeGrantRows(cur.grantRows);
    const toStore = { ...next, listProfiles, grantRows };
    try {
        localStorage.setItem(PERMS_STORAGE_KEY, JSON.stringify(toStore));
    } catch {
        /* ignore */
    }
    return { ...next, listProfiles, grantRows };
}

/** Zeilen für die Matrix-UI (inkl. leerer Startzeilen). */
export function grantRowsForUi() {
    const cur = loadPermissionsConfig();
    if (cur.grantRows && cur.grantRows.length) return normalizeGrantRows(cur.grantRows);
    const migrated = migrateLegacyPermConfig(cur);
    if (migrated.length) return migrated;
    return seedDefaultGrantRows();
}

/**
 * @param {string} typeKey
 */
export function resolveListPermProfile(typeKey) {
    const cfg = loadPermissionsConfig();
    const profiles = cfg.listProfiles || normalizeListPermProfiles(null);
    const key = String(typeKey || '').trim();
    return profiles[key] || STAMMDATEN_LIST_PERM_PROFILES[key];
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
        const withRows = {
            ...merged,
            grantRows: migrateLegacyPermConfig({ ...cur, ...merged })
        };
        savePermissionsConfig(withRows);
        return withRows;
    } catch {
        return cur;
    }
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const ROLE_ID = {
    read: SPO_ROLE.read,
    contribute: SPO_ROLE.contribute,
    fullControl: SPO_ROLE.fullControl
};

function roleDefIdForLevel(level) {
    const id = ROLE_ID[level];
    return id != null ? id : SPO_ROLE.read;
}

async function resolveGroupId(token, mailOrId) {
    const raw = String(mailOrId || '').trim();
    if (!raw) return '';
    if (GUID_RE.test(raw)) return raw;
    const esc = raw.replace(/'/g, "''");
    const filter = encodeURIComponent(
        "mail eq '" + esc + "' or mailNickname eq '" + esc + "' or displayName eq '" + esc + "'"
    );
    const data = await G().graphJson(
        'GET',
        '/groups?$filter=' + filter + '&$select=id,displayName,mail,mailNickname&$top=5',
        token,
        undefined,
        'v1.0'
    );
    const hit = ((data && data.value) || [])[0];
    return hit && hit.id ? String(hit.id) : '';
}

/**
 * @param {string} siteWebUrl
 * @param {string} spoToken
 * @param {string} digest
 * @param {string} listTitle
 * @param {string} graphToken
 * @param {{ groupId: string, groupLabel: string, level: PermLevel }[]} grants
 * @param {(msg: string) => void} write
 */
export async function applyStammdatenListGroupGrants(
    siteWebUrl,
    spoToken,
    digest,
    listTitle,
    graphToken,
    grants,
    write
) {
    await G().spoBreakListInheritance(siteWebUrl, spoToken, digest, listTitle, true);
    write('  „' + listTitle + '": Vererbung gebrochen.');

    const assignments = await G().spoListRoleAssignments(siteWebUrl, spoToken, digest, listTitle);
    let removed = 0;
    for (let i = 0; i < assignments.length; i++) {
        const a = assignments[i];
        const member = a.Member || a.member || {};
        if (!isBroadSiteAudience(member)) continue;
        const pid = member.Id != null ? member.Id : a.PrincipalId;
        try {
            await G().spoRemoveRoleAssignment(siteWebUrl, spoToken, digest, listTitle, pid);
            removed++;
        } catch {
            /* ignore */
        }
    }
    if (removed) write('  „' + listTitle + '": breite Site-Rollen entfernt (' + removed + ').');

    const list = grants && grants.length ? grants : [];
    if (!list.length) {
        write('  ! „' + listTitle + '": keine Gruppen-Zuweisungen in der Matrix.');
        return;
    }

    for (let g = 0; g < list.length; g++) {
        const grant = list[g];
        const level = grant.level;
        if (!level) continue;
        let id = grant.groupId && GUID_RE.test(grant.groupId) ? grant.groupId : '';
        if (!id && grant.groupLabel) {
            id = await resolveGroupId(graphToken, grant.groupLabel);
        }
        if (!id) {
            write(
                '  ! „' +
                    listTitle +
                    '": Gruppe nicht gewählt oder nicht gefunden (' +
                    (grant.groupLabel || 'ohne Name') +
                    ').'
            );
            continue;
        }
        try {
            const principal = await G().spoEnsureUser(siteWebUrl, spoToken, digest, entraGroupLogonName(id));
            await G().spoAddRoleAssignment(
                siteWebUrl,
                spoToken,
                digest,
                listTitle,
                principal.id,
                roleDefIdForLevel(level)
            );
            write(
                '  + „' +
                    listTitle +
                    '": ' +
                    level +
                    ' → ' +
                    (principal.title || grant.groupLabel || id)
            );
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            if (/addroleassignment:\s*500/i.test(msg) || /already|duplicate|vorhanden/i.test(msg)) {
                write('  = „' + listTitle + '": ' + (grant.groupLabel || id) + ' bereits zugewiesen.');
                continue;
            }
            write('  ! „' + listTitle + '": ' + (grant.groupLabel || id) + ' (' + level + '): ' + msg);
        }
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
    const cur = loadEffectivePermissionsConfig();
    const config = normalizePermissionsConfig({ ...cur, ...(configPatch || {}) });
    const o = listOpts && typeof listOpts === 'object' ? listOpts : {};
    if (config.skipPerms || o.skipPerms) {
        write('Berechtigungen übersprungen.');
        return { skipped: true };
    }
    const grantRows =
        configPatch && configPatch.grantRows != null
            ? normalizeGrantRows(configPatch.grantRows)
            : mergeAudienceSlotsIntoGrantRows(cur.grantRows, config);
    savePermissionsConfig({ ...config, grantRows });

    const titles = {
        schueler: String(o.schuelerTitle || 'Schülerinnen').trim() || 'Schülerinnen',
        faecher: String(o.faecherTitle || 'Fächer').trim() || 'Fächer',
        fachgruppen: String(o.fachgruppenTitle || 'Fachgruppen').trim() || 'Fachgruppen',
        arges: String(o.argesTitle || 'ARGEs').trim() || 'ARGEs',
        klassen: String(o.klassenTitle || 'Klassen').trim() || 'Klassen',
        lehrer: String(o.lehrerTitle || 'Lehrerinnen').trim() || 'Lehrerinnen'
    };

    const active = [];
    if (o.schueler) active.push('schueler');
    if (o.faecher) active.push('faecher');
    if (o.fachgruppen) active.push('fachgruppen');
    if (o.arges) active.push('arges');
    if (o.klassen) active.push('klassen');
    if (o.lehrer) active.push('lehrer');
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
        if (!listTitle) continue;
        const list = await findListByDisplayName(ctx.graphToken, siteId, listTitle);
        if (!list || !list.id) {
            write('  ! „' + listTitle + '": Liste nicht gefunden – übersprungen.');
            continue;
        }
        const displayTitle = list.displayName ? String(list.displayName) : listTitle;
        const grants = grantsForListKey(typeKey, grantRows);
        try {
            await applyStammdatenListGroupGrants(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                displayTitle,
                ctx.graphToken,
                grants,
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
