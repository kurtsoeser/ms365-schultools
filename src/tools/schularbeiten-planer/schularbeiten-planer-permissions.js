/**
 * SharePoint-Berechtigungen für Schularbeiten-Listen (Provisioning & Nachziehen).
 */
import { LIST_TITLES, LIST_KEYS, titlesForListKey } from './schularbeiten-planer-schema.js';
import { resolvePlanerList } from './schularbeiten-planer-lists.js';
import { findListByDisplayName } from './schularbeiten-planer-graph.js';
import {
    SPO_ROLE,
    entraGroupLogonName,
    isBroadSiteAudience
} from '../../shared/stammdaten-sharepoint-sync-logic.js';
import { normalizePlannerUsers } from '../freistellung-planer/freistellung-planer-direktion-users.js';
import { overlaySchoolAudienceOnPermissions } from '../../shared/school-audience-groups.js';
import { notifyAppLocalDataChanged } from '../../shared/app-local-data-notify.js';
import {
    SA_STAMMDATEN_GROUP_ROLES,
    stripStammdatenGroupFieldsFromPatch
} from '../../shared/planner-stammdaten-audience-ui.js';

export const PERMS_STORAGE_KEY = 'ms365-schularbeiten-perms-v1';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

/** @typedef {'read'|'contribute'|'edit'|'fullControl'} PermLevel */

/**
 * Rollen pro Liste (nach Vererbungsbruch).
 * @type {Record<string, { admin: PermLevel, lehrer: PermLevel|null, schueler: PermLevel|null }>}
 */
export const LIST_PERM_PROFILES = {
    schularbeiten: { admin: 'fullControl', lehrer: 'contribute', schueler: 'read' },
    regelwerk: { admin: 'fullControl', lehrer: 'read', schueler: null },
    terminfenster: { admin: 'fullControl', lehrer: 'read', schueler: null },
    fachMeta: { admin: 'fullControl', lehrer: 'read', schueler: null }
};

/** @type {Record<string, string>} */
export const LIST_TITLE_TO_PROFILE = (() => {
    const m = {
        [LIST_TITLES.schularbeiten]: 'schularbeiten',
        [LIST_TITLES.regelwerk]: 'regelwerk',
        [LIST_TITLES.terminfenster]: 'terminfenster',
        [LIST_TITLES.fachMeta]: 'fachMeta'
    };
    LIST_KEYS.forEach((key) => {
        titlesForListKey(key).forEach((title) => {
            m[title] = key;
        });
    });
    return m;
})();

export const PACKAGE_LIST_TITLES = LIST_KEYS.map((k) => LIST_TITLES[k]);

const ROLE_ID = {
    read: SPO_ROLE.read,
    contribute: SPO_ROLE.contribute,
    edit: SPO_ROLE.edit,
    fullControl: SPO_ROLE.fullControl
};

const DEFAULT_GROUPS = {
    groupAdmin: '',
    groupAdminId: '',
    groupLehrer: '',
    groupLehrerId: '',
    groupSchueler: '',
    groupSchuelerId: ''
};

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

/**
 * @param {object|null|undefined} raw
 */
export function normalizePermissionsConfig(raw) {
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

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 * @param {'groupAdmin'|'groupLehrer'|'groupSchueler'} roleKey
 */
export function groupRefFromConfig(config, roleKey) {
    const id = String(config[roleKey + 'Id'] || '').trim();
    const label = String(config[roleKey] || '').trim();
    return { id, label };
}

export function loadPermissionsConfig() {
    try {
        const raw = JSON.parse(localStorage.getItem(PERMS_STORAGE_KEY) || '{}');
        return normalizePermissionsConfig(raw);
    } catch {
        return normalizePermissionsConfig({});
    }
}

/** Stammdaten-Sammelgruppen für Lehrer/Schüler einbeziehen. */
export function loadEffectivePermissionsConfig() {
    return normalizePermissionsConfig(overlaySchoolAudienceOnPermissions(loadPermissionsConfig()));
}

/**
 * @param {object} patch
 */
export function savePermissionsConfig(patch) {
    const merged = { ...loadPermissionsConfig(), ...(patch || {}) };
    const next = normalizePermissionsConfig(
        stripStammdatenGroupFieldsFromPatch(merged, SA_STAMMDATEN_GROUP_ROLES)
    );
    try {
        localStorage.setItem(PERMS_STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    notifyAppLocalDataChanged('schularbeiten-perms');
    return next;
}

/**
 * @param {string} level
 */
export function roleDefIdForLevel(level) {
    const id = ROLE_ID[level];
    return id != null ? id : SPO_ROLE.read;
}

async function resolveGroupId(token, mailOrId) {
    const raw = String(mailOrId || '').trim();
    if (!raw) return '';
    if (/^[0-9a-f-]{36}$/i.test(raw)) return raw;
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
 * @param {ReturnType<typeof normalizePermissionsConfig>} config
 * @param {string} profileKey
 * @param {(msg: string) => void} write
 */
/**
 * @param {{ admin: PermLevel, lehrer: PermLevel|null, schueler: PermLevel|null }} profile
 */
export async function applyEntraGroupListPermissions(
    siteWebUrl,
    spoToken,
    digest,
    listTitle,
    graphToken,
    config,
    profile,
    write
) {
    if (!profile || typeof profile !== 'object') throw new Error('Berechtigungsprofil fehlt.');

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

    async function grant(roleKey, level, label) {
        if (!level) return;
        const ref = groupRefFromConfig(config, roleKey);
        let id = ref.id && GUID_RE.test(ref.id) ? ref.id : '';
        if (!id && ref.label) {
            id = await resolveGroupId(graphToken, ref.label);
        }
        if (!id) {
            write(
                '  ! „' +
                    listTitle +
                    '": Gruppe nicht gewählt oder nicht gefunden (' +
                    label +
                    '). Bitte im Gruppen-Picker eine Entra-Gruppe wählen.'
            );
            return;
        }
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
                label +
                ' → ' +
                level +
                ' (' +
                (principal.title || ref.label || id) +
                ')'
        );
    }

    if (config.groupAdmin || config.groupAdminId) await grant('groupAdmin', profile.admin, 'Verwaltung');
    if ((config.groupLehrer || config.groupLehrerId) && profile.lehrer) {
        await grant('groupLehrer', profile.lehrer, 'Lehrer');
    }
    if ((config.groupSchueler || config.groupSchuelerId) && profile.schueler) {
        await grant('groupSchueler', profile.schueler, 'Schüler');
    }

    async function grantUsers(users, level, label) {
        if (!level) return;
        for (const u of users || []) {
            const mail = String(u.mail || '').trim();
            if (!mail) continue;
            try {
                const principal = await G().spoEnsureUser(siteWebUrl, spoToken, digest, mail);
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
                        label +
                        ' (Einzelperson) → ' +
                        level +
                        ' (' +
                        (principal.title || u.displayName || mail) +
                        ')'
                );
            } catch (e) {
                write('  ! Einzelperson ' + mail + ': ' + (e && e.message ? e.message : e));
            }
        }
    }

    await grantUsers(config.adminUsers, profile.admin, 'Verwaltung');
    await grantUsers(config.lehrerUsers, profile.lehrer, 'Lehrer');
    await grantUsers(config.schuelerUsers, profile.schueler, 'Schüler');
}

export async function applySchularbeitenListPermissions(
    siteWebUrl,
    spoToken,
    digest,
    listTitle,
    graphToken,
    config,
    profileKey,
    write
) {
    const profile = LIST_PERM_PROFILES[profileKey];
    if (!profile) throw new Error('Unbekanntes Berechtigungsprofil: ' + profileKey);
    return await applyEntraGroupListPermissions(
        siteWebUrl,
        spoToken,
        digest,
        listTitle,
        graphToken,
        config,
        profile,
        write
    );
}

export async function acquireSpoContext(siteWebUrl) {
    const url = String(siteWebUrl || '').trim();
    if (!url) throw new Error('SharePoint-Website-URL fehlt.');
    const graphToken = await G().getGraphToken(SCOPES_GRAPH);
    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        throw new Error('SharePoint-Host aus URL nicht lesbar.');
    }
    const spoScope = 'https://' + host + '/Sites.FullControl.All';
    let spoToken;
    try {
        spoToken = await G().getGraphToken([spoScope]);
    } catch (e) {
        throw new Error(
            'SharePoint-Token fehlgeschlagen (Sites.FullControl.All?). ' + (e && e.message ? e.message : e)
        );
    }
    const digest = await G().getSpoRequestDigest(url, spoToken);
    return { url, graphToken, spoToken, digest };
}

/**
 * Berechtigungen auf alle vier Planer-Listen anwenden.
 * @param {string} siteWebUrl
 * @param {object} [configPatch]
 * @param {(msg: string) => void} [logFn]
 */
export async function applySchularbeitenPackagePermissions(siteWebUrl, configPatch, logFn) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const config = normalizePermissionsConfig({ ...loadPermissionsConfig(), ...(configPatch || {}) });
    if (config.skipPerms) {
        write('Berechtigungen übersprungen (skipPerms).');
        return { skipped: true };
    }
    savePermissionsConfig(config);

    write('Berechtigungen (Entra-Gruppen) …');
    const ctx = await acquireSpoContext(siteWebUrl);
    const site = await G().resolveSiteFromWebUrl(ctx.graphToken, ctx.url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID für Berechtigungen fehlt.');

    for (let i = 0; i < LIST_KEYS.length; i++) {
        const listKey = LIST_KEYS[i];
        const resolved = await resolvePlanerList(ctx.graphToken, siteId, listKey, findListByDisplayName);
        const title = resolved && resolved.displayName ? resolved.displayName : LIST_TITLES[listKey];
        const profileKey = listKey;
        if (!resolved || !resolved.list) {
            write('  ! „' + LIST_TITLES[listKey] + '": Liste auf der Site nicht gefunden.');
            continue;
        }
        try {
            await applySchularbeitenListPermissions(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                title,
                ctx.graphToken,
                config,
                profileKey,
                write
            );
        } catch (e) {
            write('  ! „' + title + '": ' + (e && e.message ? e.message : e));
        }
        await G().sleep(200);
    }
    write('Berechtigungen fertig.');
    return { skipped: false, config };
}
