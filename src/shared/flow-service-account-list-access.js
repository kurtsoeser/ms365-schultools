/**
 * Power-Automate-Technik-Konto: nach Vererbungsbruch / „breite Site-Rollen entfernt“
 * immer wieder Vollzugriff auf die Flow-Listen – sonst GetTable/Trigger 403.
 */
import { SPO_ROLE } from './stammdaten-sharepoint-sync-logic.js';
import { acquireSpoContext, roleDefIdForLevel } from '../tools/schularbeiten-planer/schularbeiten-planer-permissions.js';

/** Bekannte Setup-Keys (localStorage) mit flowServiceAccount / emailMailbox. */
export const FLOW_SETUP_STORAGE_KEYS = {
    akt: 'ms365-akt-setup-v1',
    lfr: 'ms365-lfr-setup-v1',
    freistellung: 'ms365-freistellung-setup-v1',
    srdpMail: 'ms365-srdp-mail-setup-v1'
};

/**
 * @param {Record<string, unknown>|null|undefined} cfg
 */
export function effectiveFlowServiceAccountFromCfg(cfg) {
    const c = cfg && typeof cfg === 'object' ? cfg : {};
    const explicit = String(c.flowServiceAccount || '')
        .trim()
        .toLowerCase();
    if (explicit.includes('@')) return explicit;
    const mailbox = String(c.emailMailbox || '')
        .trim()
        .toLowerCase();
    return mailbox.includes('@') ? mailbox : '';
}

/**
 * @param {string} storageKey
 * @param {Storage|{ getItem: (k: string) => string|null }} [storage]
 */
export function loadFlowServiceAccountFromStorage(storageKey, storage) {
    const store =
        storage ||
        (typeof localStorage !== 'undefined' ? localStorage : null);
    if (!store || !storageKey) return '';
    try {
        const raw = JSON.parse(store.getItem(storageKey) || '{}') || {};
        return effectiveFlowServiceAccountFromCfg(raw);
    } catch {
        return '';
    }
}

/**
 * @param {'akt'|'lfr'|'freistellung'|'srdpMail'|string} tool
 * @param {string} [explicitEmail]
 */
export function resolveFlowServiceAccount(tool, explicitEmail) {
    const explicit = String(explicitEmail || '')
        .trim()
        .toLowerCase();
    if (explicit.includes('@')) return explicit;
    const key = FLOW_SETUP_STORAGE_KEYS[tool] || tool;
    return loadFlowServiceAccountFromStorage(key);
}

/**
 * @param {string|string[]|null|undefined} emails
 * @returns {string[]}
 */
export function normalizeFlowServiceEmails(emails) {
    const seen = new Set();
    const out = [];
    const list = Array.isArray(emails) ? emails : [emails];
    list.forEach((e) => {
        const em = String(e || '')
            .trim()
            .toLowerCase();
        if (!em.includes('@') || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    });
    return out;
}

/**
 * Vollzugriff für Technik-/Flow-Konten auf Listen (idempotent).
 * @param {string} siteWebUrl
 * @param {string|string[]} emails
 * @param {string|string[]} listTitles
 * @param {(msg: string) => void} [logFn]
 * @param {{ roleLevel?: string }} [opts]
 */
export async function grantFlowServiceAccountsFullControlOnLists(
    siteWebUrl,
    emails,
    listTitles,
    logFn,
    opts
) {
    const write = typeof logFn === 'function' ? logFn : () => {};
    const uniq = normalizeFlowServiceEmails(emails);
    const titles = (Array.isArray(listTitles) ? listTitles : [listTitles])
        .map((t) => String(t || '').trim())
        .filter(Boolean);
    if (!uniq.length) {
        write('  ! Kein Power-Automate-Technik-Konto – Flow verliert nach Rechte-Schreiben ggf. den Listen-Zugriff.');
        return { skipped: true, reason: 'no-account', granted: [] };
    }
    if (!titles.length) {
        write('  ! Keine Listen für Technik-Konto-Rechte.');
        return { skipped: true, reason: 'no-lists', granted: [] };
    }
    const web = String(siteWebUrl || '').trim();
    if (!web) {
        write('  ! Site-URL fehlt für Technik-Konto-Rechte.');
        return { skipped: true, reason: 'no-site', granted: [] };
    }

    const G = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!G || typeof G.spoEnsureUser !== 'function' || typeof G.spoAddRoleAssignment !== 'function') {
        write('  ! SharePoint-Hilfen fehlen – Technik-Konto nicht zugewiesen.');
        return { skipped: true, reason: 'no-spo-api', granted: [] };
    }

    const level = (opts && opts.roleLevel) || 'fullControl';
    const roleId =
        typeof roleDefIdForLevel === 'function' ? roleDefIdForLevel(level) : SPO_ROLE.fullControl;
    const ctx = await acquireSpoContext(web);
    const granted = [];

    write('Technik-Konto (Power Automate): Vollzugriff nach Rechte-Schreiben absichern …');
    for (let t = 0; t < titles.length; t++) {
        const listTitle = titles[t];
        for (let i = 0; i < uniq.length; i++) {
            const em = uniq[i];
            try {
                const logon =
                    typeof G.membershipLogonName === 'function'
                        ? G.membershipLogonName(em)
                        : 'i:0#.f|membership|' + em;
                const principal = await G.spoEnsureUser(ctx.url, ctx.spoToken, ctx.digest, logon);
                await G.spoAddRoleAssignment(
                    ctx.url,
                    ctx.spoToken,
                    ctx.digest,
                    listTitle,
                    principal.id,
                    roleId
                );
                write(
                    '  + Flow-Technik ' +
                        em +
                        ' → ' +
                        level +
                        ' auf „' +
                        listTitle +
                        '“' +
                        (principal.title ? ' (' + principal.title + ')' : '')
                );
                granted.push({ email: em, listTitle });
            } catch (e) {
                const msg = e && e.message ? String(e.message) : String(e);
                if (/addroleassignment:\s*500/i.test(msg) || /already|duplicate|vorhanden/i.test(msg)) {
                    write('  = Flow-Technik ' + em + ' auf „' + listTitle + '“ bereits gesetzt.');
                    granted.push({ email: em, listTitle, already: true });
                    continue;
                }
                write('  ! Flow-Technik ' + em + ' auf „' + listTitle + '“: ' + msg);
            }
        }
    }
    return { skipped: false, granted, emails: uniq, titles };
}
