/**
 * tenant-settings-ui-directory-classes.js
 * @param {Record<string, any>} ctx
 */
import { groupLinkStatusView, paintStatusCell } from './tenant-workbench-chrome.js';

export function installTenantDirectoryClasses(ctx) {
    ctx.classMatchKey = function classMatchKey(row) {
        const code = ctx.normCode(row && row.code);
        const name = ctx.normCode(row && row.name);
        return code || name || '';
    }

    ctx.getClassGroupMatchByKey = function getClassGroupMatchByKey(keyRaw) {
        const key = ctx.normCode(keyRaw);
        if (!key) return null;
        const api = window.ms365AppDataV2;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const map = setup && setup.classGroupMatchByKey ? setup.classGroupMatchByKey : {};
        return map[key] || null;
    }

    ctx.patchClassGroupMatchByKey = function patchClassGroupMatchByKey(keyRaw, payload) {
        const key = ctx.normCode(keyRaw);
        if (!key) return;
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
        window.ms365AppDataV2.patchSetup({ classGroupMatchByKey: { [key]: payload } });
    }

    ctx.syncClassTeamFromGroupMatch = function syncClassTeamFromGroupMatch(row, group, mode) {
        const api = window.ms365AppDataV2;
        if (!api || typeof api.upsertClassTeam !== 'function' || !group) return;
        const gid = ctx.normStr(group.id);
        if (!gid) return;
        const nick = ctx.classRowMailNickname(row) || ctx.normStr(group.mailNickname);
        if (!nick) return;
        try {
            api.upsertClassTeam({
                stableMailNickname: nick,
                mailNickname: ctx.normStr(group.mailNickname) || nick,
                graphGroupId: gid,
                classCode: ctx.normStr(row && row.code),
                displayName: ctx.normStr(group.displayName) || ctx.normStr(row && row.name),
                abschlussJahr: ctx.normStr(row && row.year),
                mode: mode || 'matched'
            });
        } catch {
            /* classTeams optional – Register-Abgleich bleibt gültig */
        }
    }

    ctx.linkClassGroupMatch = function linkClassGroupMatch(row, group, opts) {
        const key = ctx.classMatchKey(row);
        if (!key || !group || !group.id) return;
        ctx.patchClassGroupMatchByKey(key, {
            groupId: ctx.normStr(group.id),
            displayName: ctx.normStr(group.displayName),
            mailNickname: ctx.normStr(group.mailNickname),
            mail: ctx.normStr(group.mail),
            notFound: false,
            checkedAt: new Date().toISOString()
        });
        ctx.syncClassTeamFromGroupMatch(row, group, (opts && opts.mode) || 'matched');
    }

    ctx.getArgeCatalogLink = function getArgeCatalogLink(code) {
        const api = window.ms365AppDataV2;
        if (api && typeof api.getCatalogLink === 'function') return api.getCatalogLink('arge', code);
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
        const links = (setup && setup.catalogLinks) || [];
        const c = ctx.normCode(code);
        return links.find(function (x) {
            return x && x.kind === 'arge' && ctx.normCode(x.code) === c;
        });
    }

    ctx.upsertArgeCatalogLink = function upsertArgeCatalogLink(code, group, syncStatus) {
        const api = window.ms365AppDataV2;
        if (!api || typeof api.upsertCatalogLink !== 'function') return;
        api.upsertCatalogLink({
            kind: 'arge',
            code: ctx.normCode(code),
            graphGroupId: group && group.id ? String(group.id) : '',
            displayName: group && group.displayName ? String(group.displayName) : '',
            mailNickname: group && group.mailNickname ? String(group.mailNickname) : '',
            mode: group && group.id ? 'matched' : '',
            syncStatus: syncStatus ? String(syncStatus) : ''
        });
    }

    ctx.patchArgeCatalogNotFound = function patchArgeCatalogNotFound(code) {
        ctx.upsertArgeCatalogLink(code, null, 'notFound');
    }

    ctx.buildClassGroupMatchCell = function buildClassGroupMatchCell(tr, row) {
        const td = document.createElement('td');
        td.style.fontSize = '0.88em';
        td.style.lineHeight = '1.35';
        const key = ctx.classMatchKey(row);
        if (!key) {
            paintStatusCell(td, {
                kind: 'unlinked',
                label: 'Nicht verknüpft',
                title: 'Abkürzung oder Klassenname nötig für den Abgleich'
            });
            return td;
        }
        const m = ctx.getClassGroupMatchByKey(key);
        let linkedOk = false;
        if (typeof window.ms365TenantSettingsIsClassGroupLinked === 'function') {
            const api = window.ms365AppDataV2;
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const container = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
            const classTeams =
                container && container.core && Array.isArray(container.core.classTeams) ? container.core.classTeams : [];
            linkedOk = window.ms365TenantSettingsIsClassGroupLinked(key, setup || {}, classTeams);
        } else {
            linkedOk = !!(m && m.groupId && !m.notFound);
        }
        const pendingTitle = !key
            ? 'Abkürzung oder Klassenname nötig für den Abgleich'
            : m && m.notFound
              ? 'Keine passende Klassengruppe gefunden'
              : 'Noch nicht mit einer Klassengruppe abgeglichen';
        const view = linkedOk
            ? m && (m.groupId || m.graphGroupId)
                ? groupLinkStatusView(m)
                : { kind: 'linked', label: 'Verknüpft', title: 'Status: Verknüpft' }
            : groupLinkStatusView(m && m.notFound ? m : null, pendingTitle);
        paintStatusCell(td, view, tr);
        return td;
    }

    ctx.pickClassGroupMatchFromSearch = function pickClassGroupMatchFromSearch(groups, row) {
        const key = ctx.classMatchKey(row);
        const name = ctx.normStr(row && row.name);
        const keyLc = key.toLowerCase();
        const nameLc = name.toLowerCase();
        const sanitize =
            ctx.graphApi() && typeof ctx.graphApi().sanitizeMailNickname === 'function'
                ? ctx.graphApi().sanitizeMailNickname
                : function (v) {
                      return String(v || '')
                          .replace(/[^0-9a-zA-Z]/g, '')
                          .toLowerCase();
                  };
        const keyNick = sanitize(key);
        const list = Array.isArray(groups) ? groups : [];
        for (let i = 0; i < list.length; i++) {
            const g = list[i] || {};
            const dn = ctx.normStr(g.displayName).toLowerCase();
            const nick = ctx.normStr(g.mailNickname).toLowerCase();
            const mailLocal = ctx.normStr(g.mail).split('@')[0].toLowerCase();
            if (dn === keyLc || nick === keyNick || mailLocal === keyNick) return g;
            if (nameLc && dn === nameLc) return g;
        }
        return null;
    }

    ctx.resolveClassGroupFromTenantInventory = function resolveClassGroupFromTenantInventory(row) {
        const inv = window.ms365TenantInventory;
        if (!inv || typeof inv.loadStructureState !== 'function' || typeof inv.readCache !== 'function') return null;
        const key = ctx.classMatchKey(row);
        if (!key) return null;
        const st = inv.loadStructureState();
        const rows = st && Array.isArray(st.rows) ? st.rows : [];
        const links = typeof inv.loadMatchLinks === 'function' ? inv.loadMatchLinks() || {} : {};
        const cls = rows.find(function (r) {
            return r && String(r.typ || '') === 'Klasse' && ctx.normCode(r.bezeichnung) === key;
        });
        if (!cls) return null;
        let gid = ctx.normStr(cls.tenantGroupId) || (links[String(cls.id)] && ctx.normStr(links[String(cls.id)].tenantGroupId)) || '';
        if (!gid && typeof inv.suggestGroupForUnit === 'function') gid = ctx.normStr(inv.suggestGroupForUnit(cls));
        if (!gid) return { notFound: true };
        const cache = inv.readCache() || { rows: [] };
        const hit = (cache.rows || []).find(function (g) {
            return ctx.normStr(g && g.id) === gid;
        });
        if (hit) {
            return {
                groupId: gid,
                displayName: ctx.normStr(hit.bezeichnung || hit.displayName),
                mailNickname: ctx.normStr(hit.alias || hit.mailNickname),
                mail: ctx.normStr(hit.mail),
                notFound: false,
                checkedAt: new Date().toISOString()
            };
        }
        return { groupId: gid, notFound: false, checkedAt: new Date().toISOString() };
    }

    ctx.verifyClassGroupForRow = async function verifyClassGroupForRow(row, opts) {
        const key = ctx.classMatchKey(row);
        if (!key) return { skipped: true };
        const invHit = ctx.resolveClassGroupFromTenantInventory(row);
        if (invHit) {
            if (invHit.groupId) {
                ctx.linkClassGroupMatch(
                    row,
                    {
                        id: invHit.groupId,
                        displayName: invHit.displayName,
                        mailNickname: invHit.mailNickname,
                        mail: invHit.mail
                    },
                    { mode: 'matched' }
                );
            } else {
                ctx.patchClassGroupMatchByKey(key, { notFound: true, checkedAt: new Date().toISOString() });
            }
            return { found: !!invHit.groupId, skipped: false };
        }
        const token = opts && opts.token ? opts.token : await ctx.graphApi().getGraphToken();
        const queries = [key, ctx.normStr(row && row.name)].filter(Boolean);
        const seen = new Set();
        const groups = [];
        for (let i = 0; i < queries.length; i++) {
            const q = queries[i];
            let hits = [];
            if (typeof ctx.graphApi().searchUnifiedGroups === 'function') {
                hits = await ctx.graphApi().searchUnifiedGroups(token, q);
            } else {
                const data = await ctx.graphApi().graphJson(
                    'GET',
                    '/groups?$filter=' +
                        encodeURIComponent("groupTypes/any(c:c eq 'Unified') and startswith(displayName,'" + String(q).replace(/'/g, "''") + "')") +
                        '&$select=' +
                        encodeURIComponent('id,displayName,mail,mailNickname,groupTypes') +
                        '&$top=25',
                    token,
                    undefined
                );
                hits = Array.isArray(data && data.value) ? data.value : [];
            }
            hits.forEach(function (g) {
                const id = ctx.normStr(g && g.id);
                if (!id || seen.has(id)) return;
                seen.add(id);
                groups.push(g);
            });
        }
        const match = ctx.pickClassGroupMatchFromSearch(groups, row);
        if (match && match.id) {
            ctx.linkClassGroupMatch(row, match, { mode: 'matched' });
            return { found: true, skipped: false };
        }
        ctx.patchClassGroupMatchByKey(key, { notFound: true, checkedAt: new Date().toISOString() });
        return { found: false, skipped: false };
    }

    ctx.classRowMailNickname = function classRowMailNickname(row) {
        if (!row) return '';
        let raw = '';
        if (typeof window.ms365DeriveClassStableMailNickname === 'function') {
            raw = window.ms365DeriveClassStableMailNickname(row.year || '', row.code || '', row);
        }
        if (!raw && row.stableMailNickname) {
            raw = String(row.stableMailNickname);
        }
        try {
            const api = window.ms365GraphUnifiedGroups;
            if (api && typeof api.sanitizeMailNickname === 'function') {
                return api.sanitizeMailNickname(raw);
            }
        } catch {
            /* ignore */
        }
        return String(raw || '')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9-]/g, '')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '')
            .slice(0, 60);
    }

    /**
     * @param {object} row
     * @param {{ memberEmails?: string[] }} [opts]
     */
    ctx.createClassGroupForRow = async function createClassGroupForRow(row, opts) {
        const key = ctx.classMatchKey(row);
        if (!key) return { skipped: true, reason: 'no-key' };
        const linked = ctx.getClassGroupMatchByKey(key);
        if (linked && linked.groupId) {
            return { skipped: true, reason: 'already-linked' };
        }
        const nick = ctx.classRowMailNickname(row);
        if (!nick) {
            return {
                error:
                    'Kein Mail-Nickname ableitbar – Abkürzung und Abschlussjahr prüfen (Alias-Schema in Stammdaten / Klassengruppen).'
            };
        }
        const displayName = ctx.normStr(row.name) || ctx.normStr(row.code) || nick;
        const desc =
            'Jahrgangsgruppe ' +
            displayName +
            (row.year ? ' / Abschluss ' + row.year : '') +
            ' (MS365-Schul-Tools)';
        const token = await ctx.graphApi().getGraphToken();
        const created = await ctx.graphApi().createUnifiedGroup(token, displayName, nick, desc);
        if (!created || !created.id) {
            return { error: 'Microsoft 365 hat keine Gruppe zurückgegeben.' };
        }
        const headEm = ctx.normStr(row.headEmail).toLowerCase();
        if (headEm.indexOf('@') !== -1 && typeof ctx.graphApi().ensureOwners === 'function') {
            try {
                await ctx.graphApi().ensureOwners(token, created.id, [headEm]);
            } catch {
                /* Besitzer optional */
            }
        }
        const memberEmails = Array.isArray(opts && opts.memberEmails) ? opts.memberEmails : [];
        if (memberEmails.length && typeof ctx.graphApi().syncEmailsToGroup === 'function') {
            try {
                await ctx.graphApi().syncEmailsToGroup(token, created.id, memberEmails, 'Klasse', function () {});
            } catch {
                /* Mitglieder optional */
            }
        }
        ctx.linkClassGroupMatch(row, created, { mode: 'created' });
        return { created: true, group: created, mailNickname: nick, displayName };
    }
}
