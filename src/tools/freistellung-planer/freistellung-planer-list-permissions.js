/**
 * SharePoint-Berechtigungen Freistellungsliste (Entra-Gruppen aus freistellung-planer-permissions).
 */
import {
    applyEntraGroupListPermissions,
    acquireSpoContext,
    roleDefIdForLevel
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { loadPermissionsConfig, normalizePermissionsConfig } from './freistellung-planer-permissions.js';
import {
    schuelerEntraGroupIdsForCheck,
    direktionEntraGroupIdsForCheck,
    kvEntraGroupIdsForCheck
} from './freistellung-planer-entra-role.js';
import { collectAllClassRows } from './freistellung-planer-class-context.js';
import { loadStammdaten } from './freistellung-planer-state.js';
import { normalizePlannerUsers } from './freistellung-planer-direktion-users.js';
import { entraGroupLogonName } from '../../shared/stammdaten-sharepoint-sync-logic.js';

/**
 * KV: „Gestaltung“ (enthält Listen verwalten – sonst bei ReadSecurity=2 nur eigene Elemente, auch mit „Bearbeiten“).
 * Schüler: Mitwirken – nur eigene Anträge (über ReadSecurity/WriteSecurity).
 * @type {{ admin: 'fullControl', lehrer: 'design', schueler: 'contribute' }}
 */
export const FREISTELLUNG_LIST_PROFILE = {
    admin: 'fullControl',
    lehrer: 'design',
    schueler: 'contribute'
};

/** SharePoint SP.List ReadSecurity / WriteSecurity */
export const FREISTELLUNG_LIST_ITEM_LEVEL = {
    readSecurity: 2,
    writeSecurity: 2
};

/**
 * @param {ReturnType<typeof normalizePermissionsConfig>} fr
 */
/**
 * Power-Automate-Technik-Konto: Vollzugriff auf die Freistellungsliste (Flow-Trigger & Updates).
 * @param {string} siteWebUrl
 * @param {string} listTitle
 * @param {string} serviceAccountEmail
 * @param {{ listId?: string }} [opts]
 * @param {(msg: string) => void} [logFn]
 */
export async function grantFreistellungFlowServiceAccountOnList(
    siteWebUrl,
    listTitle,
    serviceAccountEmail,
    opts,
    logFn
) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const mail = String(serviceAccountEmail || '').trim().toLowerCase();
    if (!mail) {
        write('! Technik-Konto fehlt – kein Vollzugriff für den Flow auf die Liste gesetzt.');
        return { skipped: true, reason: 'no-account' };
    }
    const title = String(listTitle || '').trim() || 'Freistellungen';
    const G = window.ms365SpoGraph;
    if (!G || typeof G.spoEnsureUser !== 'function' || typeof G.spoAddRoleAssignment !== 'function') {
        write('! SharePoint-Hilfsfunktionen fehlen – Technik-Konto nicht zugewiesen.');
        return { skipped: true, reason: 'no-spo-api' };
    }

    write('Technik-Konto für Flow: Vollzugriff auf „' + title + '“ (' + mail + ') …');
    const ctx = await acquireSpoContext(siteWebUrl);
    try {
        const principal = await G.spoEnsureUser(ctx.url, ctx.spoToken, ctx.digest, mail);
        await G.spoAddRoleAssignment(
            ctx.url,
            ctx.spoToken,
            ctx.digest,
            title,
            principal.id,
            roleDefIdForLevel(FREISTELLUNG_LIST_PROFILE.admin)
        );
        write(
            '  + Flow-Technik: Vollzugriff → ' + (principal.title || mail) + ' (Power Automate / SharePoint-Trigger)'
        );
        return { skipped: false, mail, principalId: principal.id };
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/addroleassignment:\s*500/i.test(msg) || /already|duplicate|vorhanden/i.test(msg)) {
            write('  = Flow-Technik: Vollzugriff für ' + mail + ' bereits gesetzt.');
            return { skipped: false, mail, already: true };
        }
        write('  ! Flow-Technik ' + mail + ': ' + msg);
        return { skipped: true, reason: 'error', error: e };
    }
}

export function mapFreistellungConfigForSpo(fr) {
    const c = normalizePermissionsConfig(fr);
    return {
        groupAdmin: c.groupAdmin || c.groupDirektion,
        groupAdminId: c.groupAdminId || c.groupDirektionId,
        groupLehrer: c.groupKv,
        groupLehrerId: c.groupKvId,
        groupSchueler: c.groupSchueler,
        groupSchuelerId: c.groupSchuelerId,
        skipPerms: !!c.skipPerms
    };
}

function pushKvEmail(set, raw) {
    const mail = String(raw || '').trim().toLowerCase();
    if (mail && mail.indexOf('@') !== -1) set.add(mail);
}

/**
 * Alle Klassenvorstand-Adressen für Listen-Berechtigung (Stammdaten + Setup-Katalog + KV-Einzelpersonen).
 * @param {ReturnType<typeof normalizePermissionsConfig>} [config]
 * @returns {string[]}
 */
function pushKvEmailsFromClassRow(set, row) {
    if (!row || typeof row !== 'object') return;
    pushKvEmail(set, row.headEmail);
    pushKvEmail(set, row.klassenvorstandEmail);
    pushKvEmail(set, row.kvEmail);
}

function pushKvNameFromClassRow(map, row) {
    if (!row || typeof row !== 'object') return;
    const mail = String(row.headEmail || row.kvEmail || row.klassenvorstandEmail || '')
        .trim()
        .toLowerCase();
    if (!mail || mail.indexOf('@') < 0) return;
    const name = String(
        row.headName || row.klassenvorstandName || row.kvName || row.klassenvorstand || ''
    ).trim();
    if (!map.has(mail)) {
        map.set(mail, name || mail);
    } else if (name && map.get(mail) === mail) {
        map.set(mail, name);
    }
}

/** Klassenvorstände als Planer-Einzelpersonen (Setup Schritt 5). */
export function kvUsersFromSchoolStammdaten() {
    /** @type {Map<string, string>} */
    const map = new Map();
    try {
        const st = loadStammdaten();
        const rows = collectAllClassRows({ stammdaten: { classes: (st && st.classes) || [] } });
        rows.forEach((row) => pushKvNameFromClassRow(map, row));
    } catch {
        /* ignore */
    }
    try {
        const root = typeof globalThis !== 'undefined' ? globalThis : {};
        const loadTenant =
            (typeof root.ms365TenantSettingsLoad === 'function' && root.ms365TenantSettingsLoad) ||
            (typeof root.window !== 'undefined' &&
                root.window &&
                typeof root.window.ms365TenantSettingsLoad === 'function' &&
                root.window.ms365TenantSettingsLoad);
        if (loadTenant) {
            const core = loadTenant();
            const data = (core && core.data) || core || {};
            (data.classes || []).forEach((row) => pushKvNameFromClassRow(map, row));
        }
    } catch {
        /* ignore */
    }
    return normalizePlannerUsers(
        Array.from(map.entries()).map(([mail, displayName]) => ({
            id: '',
            displayName,
            mail
        }))
    );
}

export function collectFreistellungKlassenvorstandEmails(config) {
    const raw = config && typeof config === 'object' ? config : {};
    const c = normalizePermissionsConfig(raw);
    const emails = new Set();
    for (const u of c.kvUsers || []) {
        pushKvEmail(emails, u && u.mail);
    }
    const rawCatalog = Array.isArray(raw.classCatalog) ? raw.classCatalog : [];
    for (let i = 0; i < rawCatalog.length; i++) {
        pushKvEmailsFromClassRow(emails, rawCatalog[i]);
    }
    try {
        const st = loadStammdaten();
        const classRows = collectAllClassRows({ stammdaten: { classes: (st && st.classes) || [] } });
        for (let i = 0; i < classRows.length; i++) {
            pushKvEmailsFromClassRow(emails, classRows[i]);
        }
    } catch {
        /* ignore */
    }
    try {
        if (typeof window !== 'undefined' && typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            const classes = (s && s.classes) || [];
            for (let i = 0; i < classes.length; i++) {
                pushKvEmailsFromClassRow(emails, classes[i]);
            }
        }
    } catch {
        /* ignore */
    }
    return Array.from(emails).sort();
}

/**
 * @param {Awaited<ReturnType<typeof acquireSpoContext>>} ctx
 * @param {string} listTitle
 * @param {string} mail
 * @param {string} level
 * @param {string} roleLabel
 * @param {(msg: string) => void} write
 */
async function grantEmailOnList(ctx, listTitle, mail, level, roleLabel, write) {
    const G = window.ms365SpoGraph;
    if (!G) return;
    const title = String(listTitle || '').trim();
    try {
        const principal = await G.spoEnsureUser(ctx.url, ctx.spoToken, ctx.digest, mail);
        await G.spoAddRoleAssignment(
            ctx.url,
            ctx.spoToken,
            ctx.digest,
            title,
            principal.id,
            roleDefIdForLevel(level)
        );
        write(
            '  + „' +
                title +
                '": ' +
                roleLabel +
                ' → ' +
                level +
                ' (' +
                (principal.title || mail) +
                ')'
        );
    } catch (e) {
        const msg = e && e.message ? String(e.message) : String(e);
        if (/addroleassignment:\s*500/i.test(msg) || /already|duplicate|vorhanden/i.test(msg)) {
            write('  = „' + title + '": ' + roleLabel + ' ' + mail + ' bereits zugewiesen.');
            return;
        }
        write('  ! ' + roleLabel + ' ' + mail + ': ' + msg);
    }
}

/**
 * Klassenvorstände: Gestaltung auf der Liste (alle Elemente lesbar trotz Elementregel; Planer filtert nach KV/Klasse).
 * @param {Awaited<ReturnType<typeof acquireSpoContext>>} ctx
 * @param {string} listTitle
 * @param {ReturnType<typeof normalizePermissionsConfig>} fr
 * @param {(msg: string) => void} write
 */
/**
 * @param {Awaited<ReturnType<typeof acquireSpoContext>>} ctx
 * @param {string} listTitle
 * @param {string[]} allGroupIds
 * @param {string} primaryGroupId
 * @param {string} level
 * @param {string} roleLabel
 * @param {(msg: string) => void} write
 */
async function grantExtraEntraGroupsOnList(ctx, listTitle, allGroupIds, primaryGroupId, level, roleLabel, write) {
    const G = window.ms365SpoGraph;
    if (!G) return;
    const title = String(listTitle || '').trim();
    const primary = String(primaryGroupId || '').trim().toLowerCase();
    const ids = allGroupIds || [];
    for (let i = 0; i < ids.length; i++) {
        const gid = String(ids[i] || '').trim();
        if (!gid || gid.toLowerCase() === primary) continue;
        try {
            const principal = await G.spoEnsureUser(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                entraGroupLogonName(gid)
            );
            await G.spoAddRoleAssignment(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                title,
                principal.id,
                roleDefIdForLevel(level)
            );
            write(
                '  + „' +
                    title +
                    '": ' +
                    roleLabel +
                    ' → ' +
                    level +
                    ' (' +
                    (principal.title || gid) +
                    ')'
            );
        } catch (e) {
            const msg = e && e.message ? String(e.message) : String(e);
            if (/addroleassignment:\s*500/i.test(msg) || /already|duplicate|vorhanden/i.test(msg)) {
                write('  = „' + title + '": ' + roleLabel + ' bereits zugewiesen.');
            } else {
                write('  ! ' + roleLabel + ' ' + gid + ': ' + msg);
            }
        }
    }
}

export async function grantFreistellungKlassenvorstandListAccess(ctx, listTitle, fr, write) {
    const writeFn = typeof write === 'function' ? write : () => {};
    const mails = collectFreistellungKlassenvorstandEmails(fr);
    if (!mails.length) {
        writeFn(
            '  Hinweis: Keine Klassenvorstand-E-Mails in Stammdaten/Katalog – nur Entra-Gruppe „KV“ (falls gewählt).'
        );
        return { count: 0 };
    }
    writeFn(
        'Klassenvorstände: Gestaltung auf der Liste für ' +
            mails.length +
            ' Adresse(n) (nötig wegen Elementregel „nur eigene“; Planer filtert nach Klasse/KV) …'
    );
    for (let i = 0; i < mails.length; i++) {
        await grantEmailOnList(
            ctx,
            listTitle,
            mails[i],
            FREISTELLUNG_LIST_PROFILE.lehrer,
            'Klassenvorstand',
            writeFn
        );
    }
    return { count: mails.length };
}

/**
 * @param {string} siteWebUrl
 * @param {string} listTitle
 * @param {object} [configPatch]
 * @param {(msg: string) => void} [logFn]
 */
export async function applyFreistellungListPermissions(siteWebUrl, listTitle, configPatch, logFn) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const patch = configPatch || {};
    const flowServiceAccount = String(patch.flowServiceAccount || '').trim().toLowerCase();
    const fr = normalizePermissionsConfig({ ...loadPermissionsConfig(), ...patch });
    const title = String(listTitle || '').trim() || 'Freistellungen';
    const listId = String(patch.listId || fr.listId || '').trim();

    if (fr.skipPerms) {
        write('Entra-Gruppen-Berechtigungen übersprungen (Checkbox aktiv).');
        const ctxKv = await acquireSpoContext(siteWebUrl);
        await grantFreistellungKlassenvorstandListAccess(ctxKv, title, fr, write);
        const flowOnly = await grantFreistellungFlowServiceAccountOnList(
            siteWebUrl,
            title,
            flowServiceAccount,
            { listId },
            write
        );
        return { skipped: true, flowServiceAccount: flowOnly };
    }
    const direktionGroupIds = direktionEntraGroupIdsForCheck(fr);
    const kvGroupIds = kvEntraGroupIdsForCheck(fr);
    const schuelerGroupIds = schuelerEntraGroupIdsForCheck(fr);
    if (!direktionGroupIds.length && !kvGroupIds.length && !schuelerGroupIds.length) {
        write('! Berechtigungen: keine Entra-Gruppen gewählt – bitte im Setup eintragen.');
        const ctxKv = await acquireSpoContext(siteWebUrl);
        await grantFreistellungKlassenvorstandListAccess(ctxKv, title, fr, write);
        const flowOnly = await grantFreistellungFlowServiceAccountOnList(
            siteWebUrl,
            title,
            flowServiceAccount,
            { listId },
            write
        );
        return { skipped: true, reason: 'no-groups', flowServiceAccount: flowOnly };
    }

    const mapped = mapFreistellungConfigForSpo(fr);
    if (!mapped.groupAdminId && direktionGroupIds[0]) {
        mapped.groupAdminId = direktionGroupIds[0];
    }
    if (!mapped.groupLehrerId && kvGroupIds[0]) {
        mapped.groupLehrerId = kvGroupIds[0];
    }
    if (!mapped.groupSchuelerId && schuelerGroupIds[0]) {
        mapped.groupSchuelerId = schuelerGroupIds[0];
    }
    write('Berechtigungen auf Liste „' + title + '“ …');
    const ctx = await acquireSpoContext(siteWebUrl);
    await applyEntraGroupListPermissions(
        ctx.url,
        ctx.spoToken,
        ctx.digest,
        title,
        ctx.graphToken,
        mapped,
        FREISTELLUNG_LIST_PROFILE,
        write
    );

    const G = window.ms365SpoGraph;
    if (G && !fr.skipPerms) {
        await grantExtraEntraGroupsOnList(
            ctx,
            title,
            direktionGroupIds,
            mapped.groupAdminId,
            FREISTELLUNG_LIST_PROFILE.admin,
            'Direktion (weitere Gruppe)',
            write
        );
        await grantExtraEntraGroupsOnList(
            ctx,
            title,
            kvGroupIds,
            mapped.groupLehrerId,
            FREISTELLUNG_LIST_PROFILE.lehrer,
            'KV (weitere Gruppe)',
            write
        );
        await grantExtraEntraGroupsOnList(
            ctx,
            title,
            schuelerGroupIds,
            mapped.groupSchuelerId,
            FREISTELLUNG_LIST_PROFILE.schueler,
            'Schüler (weitere Gruppe)',
            write
        );
        async function grantUsers(users, level, roleLabel) {
            for (const u of users || []) {
                const mail = String(u.mail || '').trim();
                if (!mail) continue;
                await grantEmailOnList(ctx, title, mail, level, roleLabel, write);
            }
        }
        await grantUsers(fr.direktionUsers, FREISTELLUNG_LIST_PROFILE.admin, 'Direktion');
        await grantUsers(fr.schuelerUsers, FREISTELLUNG_LIST_PROFILE.schueler, 'Schüler');
    }

    await grantFreistellungKlassenvorstandListAccess(ctx, title, fr, write);

    if (G && typeof G.spoSetListItemLevelPermissions === 'function') {
        try {
            await G.spoSetListItemLevelPermissions(
                ctx.url,
                ctx.spoToken,
                ctx.digest,
                title,
                FREISTELLUNG_LIST_ITEM_LEVEL.readSecurity,
                FREISTELLUNG_LIST_ITEM_LEVEL.writeSecurity,
                listId || undefined
            );
            let verified = false;
            if (typeof G.spoGetListItemLevelPermissions === 'function') {
                const cur = await G.spoGetListItemLevelPermissions(
                    ctx.url,
                    ctx.spoToken,
                    ctx.digest,
                    title,
                    listId || undefined
                );
                verified =
                    cur.readSecurity === FREISTELLUNG_LIST_ITEM_LEVEL.readSecurity &&
                    cur.writeSecurity === FREISTELLUNG_LIST_ITEM_LEVEL.writeSecurity;
                if (verified) {
                    write(
                        '  „' +
                            title +
                            '": Elementberechtigungen aktiv (Lesen/Schreiben nur eigene Elemente für Mitwirkende – Schüler).'
                    );
                } else {
                    write(
                        '  ! Elementberechtigungen nicht übernommen (ReadSecurity=' +
                            cur.readSecurity +
                            ', WriteSecurity=' +
                            cur.writeSecurity +
                            '). In SharePoint manuell: Listen-Einstellungen → Erweitert → „Elemente lesen: nur eigene“.'
                    );
                }
            } else {
                write(
                    '  „' +
                        title +
                        '": Elementberechtigungen gesetzt (Schüler nur eigene Anträge – bitte als Schüler testen).'
                );
            }
        } catch (e) {
            write(
                '  ! Elementberechtigungen: ' + (e && e.message ? e.message : e)
            );
        }
    }

    await grantFreistellungFlowServiceAccountOnList(
        siteWebUrl,
        title,
        flowServiceAccount,
        { listId },
        write
    );

    write('Berechtigungen fertig.');
    return { skipped: false };
}
