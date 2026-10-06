/**
 * SharePoint-Berechtigungen Freistellungsliste (Entra-Gruppen aus freistellung-planer-permissions).
 */
import {
    applyEntraGroupListPermissions,
    acquireSpoContext,
    roleDefIdForLevel
} from '../schularbeiten-planer/schularbeiten-planer-permissions.js';
import { loadPermissionsConfig, normalizePermissionsConfig } from './freistellung-planer-permissions.js';

/**
 * KV: „Bearbeiten“ (sieht alle Elemente trotz Elementberechtigung „nur eigene“ für Mitwirkende).
 * Schüler: Mitwirken – nur eigene Anträge (über ReadSecurity/WriteSecurity).
 * @type {{ admin: 'fullControl', lehrer: 'edit', schueler: 'contribute' }}
 */
export const FREISTELLUNG_LIST_PROFILE = {
    admin: 'fullControl',
    lehrer: 'edit',
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
        groupAdmin: c.groupDirektion,
        groupAdminId: c.groupDirektionId,
        groupLehrer: c.groupKv,
        groupLehrerId: c.groupKvId,
        groupSchueler: c.groupSchueler,
        groupSchuelerId: c.groupSchuelerId,
        skipPerms: !!c.skipPerms
    };
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
        const flowOnly = await grantFreistellungFlowServiceAccountOnList(
            siteWebUrl,
            title,
            flowServiceAccount,
            { listId },
            write
        );
        return { skipped: true, flowServiceAccount: flowOnly };
    }
    if (!fr.groupDirektionId && !fr.groupKvId && !fr.groupSchuelerId) {
        write('! Berechtigungen: keine Entra-Gruppen gewählt – bitte im Setup eintragen.');
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
        async function grantUsers(users, level, roleLabel) {
            for (const u of users || []) {
                const mail = String(u.mail || '').trim();
                if (!mail) continue;
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
                            (principal.title || u.displayName || mail) +
                            ')'
                    );
                } catch (e) {
                    write('  ! Einzelperson ' + mail + ': ' + (e && e.message ? e.message : e));
                }
            }
        }
        await grantUsers(fr.direktionUsers, FREISTELLUNG_LIST_PROFILE.admin, 'Direktion');
        await grantUsers(fr.kvUsers, FREISTELLUNG_LIST_PROFILE.lehrer, 'KV');
        await grantUsers(fr.schuelerUsers, FREISTELLUNG_LIST_PROFILE.schueler, 'Schüler');
    }

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
